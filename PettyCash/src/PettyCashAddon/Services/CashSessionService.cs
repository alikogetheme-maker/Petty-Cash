using System;
using System.Collections.Generic;
using System.Globalization;
using SAPbobsCOM;
using PettyCashAddon.Core;
using PettyCashAddon.Models;

namespace PettyCashAddon.Services
{
    internal class SessionRow
    {
        public string Code;
        public DateTime CashDate;
        public Shift Shift;
        public string Cashier;
        public double OpenBal;
        public double TheoBal;
        public double CountBal;
        public double Diff;
        public SessionStatus Status;
        public DateTime? ClosedAt;
    }

    internal class TransactionRow
    {
        public string Code;
        public string SessionCode;
        public string Time;
        public Direction Direction;
        public string TransTypeCode;
        public double Amount;
        public string CardCode;
        public string Description;
        public string JeDocEntry;
    }

    internal class TransactionType
    {
        public string Code;
        public string Name;
        public Direction Direction;
        public string GlAccount;
    }

    internal class CashSettings
    {
        public string CashAccount;
        public string DiffAccount;
    }

    /// <summary>
    /// Toute la logique métier de la caisse : ouverture de session, ajout
    /// de transaction (avec comptabilisation immédiate), clôture avec
    /// calcul d'écart. C'est le fichier qui fait référence pour comprendre
    /// le comportement fonctionnel de l'add-on.
    /// </summary>
    internal static class CashSessionService
    {
        // ---------- Paramétrage ----------

        public static CashSettings GetSettings()
        {
            UserTable table = (UserTable)DiCompany.Instance.GetBusinessObject(BoObjectTypes.oUserTable);
            try
            {
                table.TableName = Db.SettingsTable;
                if (!table.GetByKey(Db.SettingsCode))
                    throw new InvalidOperationException("Paramétrage caisse introuvable (table @" + Db.SettingsTable + ").");

                string cashAcct = Convert.ToString(table.UserFields.Fields.Item(Db.F_CashAcct).Value).Trim();
                string diffAcct = Convert.ToString(table.UserFields.Fields.Item(Db.F_DiffAcct).Value).Trim();

                if (string.IsNullOrEmpty(cashAcct) || string.IsNullOrEmpty(diffAcct))
                    throw new InvalidOperationException("Renseignez les comptes G/L (U_CashAcct, U_DiffAcct) dans @" + Db.SettingsTable + " avant d'utiliser la caisse.");

                return new CashSettings { CashAccount = cashAcct, DiffAccount = diffAcct };
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
            }
        }

        public static List<TransactionType> GetTransactionTypes(Direction? filterDirection = null)
        {
            var result = new List<TransactionType>();
            Recordset rs = (Recordset)DiCompany.Instance.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                string sql = "SELECT \"Code\", \"Name\", \"" + Db.F_TType_Dir + "\", \"" + Db.F_TType_GLAcct + "\" FROM \"@" + Db.TTypeTable + "\"";
                if (filterDirection.HasValue)
                    sql += " WHERE \"" + Db.F_TType_Dir + "\" = '" + EnumCodes.ToCode(filterDirection.Value) + "'";
                sql += " ORDER BY \"Name\"";

                rs.DoQuery(sql);
                while (!rs.EoF)
                {
                    result.Add(new TransactionType
                    {
                        Code = Convert.ToString(rs.Fields.Item(0).Value),
                        Name = Convert.ToString(rs.Fields.Item(1).Value),
                        Direction = EnumCodes.DirectionFromCode(Convert.ToString(rs.Fields.Item(2).Value)),
                        GlAccount = Convert.ToString(rs.Fields.Item(3).Value)
                    });
                    rs.MoveNext();
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
            return result;
        }

        // ---------- Sessions ----------

        public static SessionRow GetOpenSession()
        {
            Recordset rs = (Recordset)DiCompany.Instance.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                string sql = "SELECT \"Code\" FROM \"@" + Db.SessionTable + "\" WHERE \"" + Db.F_Session_Status + "\" = 'O'";
                rs.DoQuery(sql);
                if (rs.EoF)
                    return null;
                string code = Convert.ToString(rs.Fields.Item(0).Value);
                return LoadSession(code);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }

        public static SessionRow LoadSession(string code)
        {
            UserTable table = (UserTable)DiCompany.Instance.GetBusinessObject(BoObjectTypes.oUserTable);
            try
            {
                table.TableName = Db.SessionTable;
                if (!table.GetByKey(code))
                    return null;

                return MapSession(table);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
            }
        }

        private static SessionRow MapSession(UserTable table)
        {
            object closedAtVal = table.UserFields.Fields.Item(Db.F_Session_ClosedAt).Value;
            DateTime? closedAt = null;
            if (closedAtVal != null && closedAtVal is DateTime dt && dt != DateTime.MinValue)
                closedAt = dt;

            return new SessionRow
            {
                Code = table.Code,
                CashDate = Convert.ToDateTime(table.UserFields.Fields.Item(Db.F_Session_Date).Value),
                Shift = EnumCodes.ShiftFromCode(Convert.ToString(table.UserFields.Fields.Item(Db.F_Session_Shift).Value)),
                Cashier = Convert.ToString(table.UserFields.Fields.Item(Db.F_Session_Cashier).Value),
                OpenBal = Convert.ToDouble(table.UserFields.Fields.Item(Db.F_Session_OpenBal).Value),
                TheoBal = Convert.ToDouble(table.UserFields.Fields.Item(Db.F_Session_TheoBal).Value),
                CountBal = Convert.ToDouble(table.UserFields.Fields.Item(Db.F_Session_CountBal).Value),
                Diff = Convert.ToDouble(table.UserFields.Fields.Item(Db.F_Session_Diff).Value),
                Status = EnumCodes.StatusFromCode(Convert.ToString(table.UserFields.Fields.Item(Db.F_Session_Status).Value)),
                ClosedAt = closedAt
            };
        }

        /// <summary>
        /// Ouvre une nouvelle session. Le solde d'ouverture reprend
        /// automatiquement le solde compté (ou théorique à défaut) de la
        /// dernière session clôturée. Si aucune session n'existe encore,
        /// openingBalanceIfFirst est utilisé (saisie manuelle initiale).
        /// </summary>
        public static SessionRow OpenSession(Shift shift, string cashier, double openingBalanceIfFirst)
        {
            if (GetOpenSession() != null)
                throw new InvalidOperationException("Une session de caisse est déjà ouverte. Clôturez-la avant d'en ouvrir une nouvelle.");

            double openingBalance = GetLastClosingBalance() ?? openingBalanceIfFirst;

            string code = DateTime.Now.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture);

            UserTable table = (UserTable)DiCompany.Instance.GetBusinessObject(BoObjectTypes.oUserTable);
            try
            {
                table.TableName = Db.SessionTable;
                table.Code = code;
                table.Name = DateTime.Today.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) + " - " + shift;
                table.UserFields.Fields.Item(Db.F_Session_Date).Value = DateTime.Today;
                table.UserFields.Fields.Item(Db.F_Session_Shift).Value = EnumCodes.ToCode(shift);
                table.UserFields.Fields.Item(Db.F_Session_Cashier).Value = cashier;
                table.UserFields.Fields.Item(Db.F_Session_OpenBal).Value = openingBalance;
                table.UserFields.Fields.Item(Db.F_Session_TheoBal).Value = openingBalance;
                table.UserFields.Fields.Item(Db.F_Session_CountBal).Value = 0;
                table.UserFields.Fields.Item(Db.F_Session_Diff).Value = 0;
                table.UserFields.Fields.Item(Db.F_Session_Status).Value = "O";

                int rc = table.Add();
                DiCompany.ThrowIfError(rc, "Ouverture de la session de caisse");
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
            }

            return LoadSession(code);
        }

        private static double? GetLastClosingBalance()
        {
            Recordset rs = (Recordset)DiCompany.Instance.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                string sql = "SELECT TOP 1 \"" + Db.F_Session_CountBal + "\", \"" + Db.F_Session_TheoBal + "\" " +
                             "FROM \"@" + Db.SessionTable + "\" WHERE \"" + Db.F_Session_Status + "\" = 'C' " +
                             "ORDER BY \"" + Db.F_Session_ClosedAt + "\" DESC";
                rs.DoQuery(sql);
                if (rs.EoF)
                    return null;

                double countBal = Convert.ToDouble(rs.Fields.Item(0).Value);
                double theoBal = Convert.ToDouble(rs.Fields.Item(1).Value);
                return countBal != 0 ? countBal : theoBal;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }

        /// <summary>
        /// Ajoute une transaction à la session ouverte : comptabilise
        /// immédiatement l'écriture au journal (Caisse vs compte G/L du
        /// type d'opération) puis enregistre la ligne et met à jour le
        /// solde théorique de la session.
        /// </summary>
        public static void AddTransaction(SessionRow session, Direction direction, TransactionType type, double amount, string cardCode, string description)
        {
            if (session.Status != SessionStatus.Open)
                throw new InvalidOperationException("Cette session est clôturée, impossible d'ajouter une transaction.");
            if (amount <= 0)
                throw new ArgumentException("Le montant doit être positif.");

            CashSettings settings = GetSettings();

            string jeDocEntry = PostJournalEntry(
                direction == Direction.Recette ? settings.CashAccount : type.GlAccount,
                direction == Direction.Recette ? type.GlAccount : settings.CashAccount,
                amount,
                "Caisse - " + type.Name + " - " + description);

            string code = DateTime.Now.ToString("yyyyMMddHHmmssfff", CultureInfo.InvariantCulture);
            UserTable table = (UserTable)DiCompany.Instance.GetBusinessObject(BoObjectTypes.oUserTable);
            try
            {
                table.TableName = Db.TransTable;
                table.Code = code;
                table.Name = type.Name + " " + amount.ToString(CultureInfo.InvariantCulture);
                table.UserFields.Fields.Item(Db.F_Trans_Session).Value = session.Code;
                table.UserFields.Fields.Item(Db.F_Trans_Time).Value = DateTime.Now.ToString("HH:mm:ss");
                table.UserFields.Fields.Item(Db.F_Trans_Dir).Value = EnumCodes.ToCode(direction);
                table.UserFields.Fields.Item(Db.F_Trans_TType).Value = type.Code;
                table.UserFields.Fields.Item(Db.F_Trans_Amount).Value = amount;
                table.UserFields.Fields.Item(Db.F_Trans_CardCode).Value = cardCode ?? "";
                table.UserFields.Fields.Item(Db.F_Trans_Descript).Value = description ?? "";
                table.UserFields.Fields.Item(Db.F_Trans_JE).Value = jeDocEntry;

                int rc = table.Add();
                DiCompany.ThrowIfError(rc, "Enregistrement de la transaction de caisse");
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
            }

            double delta = direction == Direction.Recette ? amount : -amount;
            UpdateTheoBalance(session.Code, session.TheoBal + delta);
        }

        private static void UpdateTheoBalance(string sessionCode, double newTheoBalance)
        {
            UserTable table = (UserTable)DiCompany.Instance.GetBusinessObject(BoObjectTypes.oUserTable);
            try
            {
                table.TableName = Db.SessionTable;
                table.GetByKey(sessionCode);
                table.UserFields.Fields.Item(Db.F_Session_TheoBal).Value = newTheoBalance;
                int rc = table.Update();
                DiCompany.ThrowIfError(rc, "Mise à jour du solde théorique de la session");
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
            }
        }

        /// <summary>
        /// Clôture la session : enregistre le solde compté, calcule
        /// l'écart et, s'il est non nul, le comptabilise sur le compte
        /// d'écart dédié.
        /// </summary>
        public static SessionRow CloseSession(SessionRow session, double countedBalance)
        {
            if (session.Status != SessionStatus.Open)
                throw new InvalidOperationException("Cette session est déjà clôturée.");

            double diff = countedBalance - session.TheoBal;

            if (Math.Abs(diff) > 0.0001)
            {
                CashSettings settings = GetSettings();
                // écart positif (compté > théorique) => excédent => on
                // ajoute l'excédent en caisse en le créditant depuis le
                // compte d'écart ; écart négatif => manquant => inverse.
                if (diff > 0)
                    PostJournalEntry(settings.CashAccount, settings.DiffAccount, diff, "Caisse - Écart de clôture (excédent) - session " + session.Code);
                else
                    PostJournalEntry(settings.DiffAccount, settings.CashAccount, -diff, "Caisse - Écart de clôture (manquant) - session " + session.Code);
            }

            UserTable table = (UserTable)DiCompany.Instance.GetBusinessObject(BoObjectTypes.oUserTable);
            try
            {
                table.TableName = Db.SessionTable;
                table.GetByKey(session.Code);
                table.UserFields.Fields.Item(Db.F_Session_CountBal).Value = countedBalance;
                table.UserFields.Fields.Item(Db.F_Session_Diff).Value = diff;
                table.UserFields.Fields.Item(Db.F_Session_Status).Value = "C";
                table.UserFields.Fields.Item(Db.F_Session_ClosedAt).Value = DateTime.Now;
                int rc = table.Update();
                DiCompany.ThrowIfError(rc, "Clôture de la session de caisse");
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
            }

            return LoadSession(session.Code);
        }

        // ---------- Transactions d'une session ----------

        public static List<TransactionRow> GetTransactions(string sessionCode)
        {
            var result = new List<TransactionRow>();
            Recordset rs = (Recordset)DiCompany.Instance.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                string sql = "SELECT \"Code\", \"" + Db.F_Trans_Session + "\", \"" + Db.F_Trans_Time + "\", \"" +
                             Db.F_Trans_Dir + "\", \"" + Db.F_Trans_TType + "\", \"" + Db.F_Trans_Amount + "\", \"" +
                             Db.F_Trans_CardCode + "\", \"" + Db.F_Trans_Descript + "\", \"" + Db.F_Trans_JE + "\" " +
                             "FROM \"@" + Db.TransTable + "\" WHERE \"" + Db.F_Trans_Session + "\" = '" + sessionCode + "' " +
                             "ORDER BY \"" + Db.F_Trans_Time + "\"";
                rs.DoQuery(sql);
                while (!rs.EoF)
                {
                    result.Add(new TransactionRow
                    {
                        Code = Convert.ToString(rs.Fields.Item(0).Value),
                        SessionCode = Convert.ToString(rs.Fields.Item(1).Value),
                        Time = Convert.ToString(rs.Fields.Item(2).Value),
                        Direction = EnumCodes.DirectionFromCode(Convert.ToString(rs.Fields.Item(3).Value)),
                        TransTypeCode = Convert.ToString(rs.Fields.Item(4).Value),
                        Amount = Convert.ToDouble(rs.Fields.Item(5).Value),
                        CardCode = Convert.ToString(rs.Fields.Item(6).Value),
                        Description = Convert.ToString(rs.Fields.Item(7).Value),
                        JeDocEntry = Convert.ToString(rs.Fields.Item(8).Value)
                    });
                    rs.MoveNext();
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
            return result;
        }

        // ---------- Rapport ----------

        public static List<SessionRow> GetSessions(DateTime from, DateTime to, Shift? shiftFilter)
        {
            var result = new List<SessionRow>();
            Recordset rs = (Recordset)DiCompany.Instance.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                string sql = "SELECT \"Code\" FROM \"@" + Db.SessionTable + "\" " +
                             "WHERE \"" + Db.F_Session_Date + "\" >= '" + from.ToString("yyyyMMdd") + "' " +
                             "AND \"" + Db.F_Session_Date + "\" <= '" + to.ToString("yyyyMMdd") + "'";
                if (shiftFilter.HasValue)
                    sql += " AND \"" + Db.F_Session_Shift + "\" = '" + EnumCodes.ToCode(shiftFilter.Value) + "'";
                sql += " ORDER BY \"" + Db.F_Session_Date + "\" DESC, \"" + Db.F_Session_Shift + "\"";

                rs.DoQuery(sql);
                var codes = new List<string>();
                while (!rs.EoF)
                {
                    codes.Add(Convert.ToString(rs.Fields.Item(0).Value));
                    rs.MoveNext();
                }

                foreach (var code in codes)
                    result.Add(LoadSession(code));
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
            return result;
        }

        // ---------- Comptabilisation ----------

        private static string PostJournalEntry(string debitAccount, string creditAccount, double amount, string memo)
        {
            JournalEntries je = (JournalEntries)DiCompany.Instance.GetBusinessObject(BoObjectTypes.oJournalEntries);
            try
            {
                je.Memo = memo.Length > 100 ? memo.Substring(0, 100) : memo;
                je.ReferenceDate = DateTime.Today;
                je.TaxDate = DateTime.Today;

                je.Lines.AccountCode = debitAccount;
                je.Lines.Debit = amount;
                je.Lines.Add();

                je.Lines.AccountCode = creditAccount;
                je.Lines.Credit = amount;

                int rc = je.Add();
                DiCompany.ThrowIfError(rc, "Comptabilisation de l'écriture de caisse");

                return DiCompany.Instance.GetNewObjectKey();
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(je);
            }
        }
    }
}
