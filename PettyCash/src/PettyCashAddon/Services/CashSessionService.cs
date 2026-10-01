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
        public string TransTypeName;
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
    ///
    /// Règles :
    /// - une seule session ouverte à la fois ;
    /// - solde d'ouverture = solde compté de la dernière session clôturée
    ///   (ou, pour la toute première session, solde du compte G/L Caisse) ;
    /// - solde théorique = ouverture + recettes - dépenses, recalculé depuis
    ///   la base à chaque opération ;
    /// - une dépense ne peut pas rendre le solde théorique négatif ;
    /// - écriture comptable + ligne de caisse + mise à jour de session sont
    ///   faites dans une même transaction DI API (tout ou rien).
    /// </summary>
    internal static class CashSessionService
    {
        // Longueurs maximales des champs SAP / UDF alimentés
        private const int MaxJeMemo = 50;        // OJDT.Memo, JDT1.LineMemo
        private const int MaxUdtName = 30;
        private const int MaxCashier = 50;
        private const int MaxCardCode = 15;
        private const int MaxDescription = 100;

        // ---------- Paramétrage ----------

        public static CashSettings GetSettings()
        {
            UserTable table = DiCompany.Instance.UserTables.Item(Db.SettingsTable);
            try
            {
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
            Recordset rs = NewRecordset();
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
                        GlAccount = Convert.ToString(rs.Fields.Item(3).Value).Trim()
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
            Recordset rs = NewRecordset();
            try
            {
                string sql = "SELECT \"Code\" FROM \"@" + Db.SessionTable + "\" WHERE \"" + Db.F_Session_Status + "\" = 'O' ORDER BY \"Code\" DESC";
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
            UserTable table = DiCompany.Instance.UserTables.Item(Db.SessionTable);
            try
            {
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
            // Un champ date vide est renvoyé par la DI API comme 30/12/1899, pas comme null
            object closedAtVal = table.UserFields.Fields.Item(Db.F_Session_ClosedAt).Value;
            DateTime? closedAt = null;
            if (closedAtVal is DateTime dt && dt.Year > 1900)
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
        /// Solde d'ouverture qui sera repris à la prochaine ouverture de
        /// session : solde compté de la dernière session clôturée, ou, s'il
        /// n'y en a encore aucune, solde actuel du compte G/L Caisse (le
        /// fonds de caisse initial doit donc y avoir été comptabilisé).
        /// </summary>
        public static double GetNextOpeningBalance()
        {
            double? last = GetLastClosingBalance();
            if (last.HasValue)
                return last.Value;

            return GetAccountBalance(GetSettings().CashAccount);
        }

        /// <summary>
        /// Ouvre une nouvelle session avec le solde d'ouverture calculé par
        /// GetNextOpeningBalance.
        /// </summary>
        public static SessionRow OpenSession(Shift shift, string cashier)
        {
            if (string.IsNullOrWhiteSpace(cashier))
                throw new ArgumentException("Renseignez le nom du caissier.");
            if (GetOpenSession() != null)
                throw new InvalidOperationException("Une session de caisse est déjà ouverte. Clôturez-la avant d'en ouvrir une nouvelle.");

            double openingBalance = Round(GetNextOpeningBalance());

            string code = DateTime.Now.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture);

            UserTable table = DiCompany.Instance.UserTables.Item(Db.SessionTable);
            try
            {
                table.Code = code;
                table.Name = Truncate(DateTime.Today.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) + " " + EnumCodes.ToCode(shift) + " " + cashier.Trim(), MaxUdtName);
                table.UserFields.Fields.Item(Db.F_Session_Date).Value = DateTime.Today;
                table.UserFields.Fields.Item(Db.F_Session_Shift).Value = EnumCodes.ToCode(shift);
                table.UserFields.Fields.Item(Db.F_Session_Cashier).Value = Truncate(cashier.Trim(), MaxCashier);
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
            Recordset rs = NewRecordset();
            try
            {
                // Le Code est l'horodatage d'ouverture (yyyyMMddHHmmss) et une seule
                // session est ouverte à la fois : trier sur le Code donne bien la
                // dernière session. U_ClosedAt ne porte que la date (pas l'heure),
                // il ne permet pas de départager plusieurs quarts d'une même journée.
                string sql = "SELECT TOP 1 \"" + Db.F_Session_CountBal + "\" " +
                             "FROM \"@" + Db.SessionTable + "\" WHERE \"" + Db.F_Session_Status + "\" = 'C' " +
                             "ORDER BY \"Code\" DESC";
                rs.DoQuery(sql);
                if (rs.EoF)
                    return null;

                // Après clôture, l'écart a été comptabilisé : le solde réel de la
                // caisse est le solde compté, y compris lorsqu'il vaut 0.
                return Convert.ToDouble(rs.Fields.Item(0).Value);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }

        /// <summary>
        /// Solde théorique recalculé depuis la base : ouverture + recettes - dépenses.
        /// </summary>
        private static double ComputeTheoBalance(SessionRow session)
        {
            Recordset rs = NewRecordset();
            try
            {
                string sql = "SELECT COALESCE(SUM(CASE WHEN \"" + Db.F_Trans_Dir + "\" = 'R' THEN \"" + Db.F_Trans_Amount + "\" " +
                             "ELSE -\"" + Db.F_Trans_Amount + "\" END), 0) " +
                             "FROM \"@" + Db.TransTable + "\" WHERE \"" + Db.F_Trans_Session + "\" = '" + Sql(session.Code) + "'";
                rs.DoQuery(sql);
                return Round(session.OpenBal + Convert.ToDouble(rs.Fields.Item(0).Value));
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }

        /// <summary>
        /// Ajoute une transaction à la session ouverte : comptabilise
        /// l'écriture au journal (Caisse vs compte G/L du type d'opération),
        /// enregistre la ligne et met à jour le solde théorique, le tout dans
        /// une même transaction DI API.
        /// </summary>
        public static void AddTransaction(string sessionCode, Direction direction, TransactionType type, double amount, string cardCode, string description)
        {
            SessionRow session = LoadSession(sessionCode);
            if (session == null)
                throw new InvalidOperationException("Session de caisse introuvable.");
            if (session.Status != SessionStatus.Open)
                throw new InvalidOperationException("Cette session est clôturée, impossible d'ajouter une transaction.");
            if (type == null)
                throw new ArgumentException("Sélectionnez un type d'opération.");
            if (type.Direction != direction)
                throw new ArgumentException("Le type d'opération « " + type.Name + " » ne correspond pas au sens choisi.");
            if (string.IsNullOrEmpty(type.GlAccount))
                throw new InvalidOperationException("Aucun compte G/L n'est paramétré pour le type d'opération « " + type.Name + " ».");

            amount = Round(amount);
            if (amount <= 0)
                throw new ArgumentException("Le montant doit être positif.");

            cardCode = (cardCode ?? "").Trim();
            if (cardCode.Length > 0 && !BusinessPartnerExists(cardCode))
                throw new ArgumentException("Le tiers « " + cardCode + " » n'existe pas dans SAP.");

            description = Truncate((description ?? "").Trim(), MaxDescription);

            CashSettings settings = GetSettings();

            double theoBefore = ComputeTheoBalance(session);
            double theoAfter = Round(theoBefore + (direction == Direction.Recette ? amount : -amount));
            if (theoAfter < 0)
                throw new InvalidOperationException("Dépense refusée : le solde théorique de la caisse (" +
                    FormatAmount(theoBefore) + ") est insuffisant pour sortir " + FormatAmount(amount) + ".");

            string memo = "Caisse " + type.Name + (description.Length > 0 ? " - " + description : "");

            Company company = DiCompany.Instance;
            company.StartTransaction();
            try
            {
                string jeDocEntry = PostJournalEntry(
                    direction == Direction.Recette ? settings.CashAccount : type.GlAccount,
                    direction == Direction.Recette ? type.GlAccount : settings.CashAccount,
                    amount, memo, session.Code);

                string code = DateTime.Now.ToString("yyyyMMddHHmmssfff", CultureInfo.InvariantCulture);
                UserTable table = company.UserTables.Item(Db.TransTable);
                try
                {
                    table.Code = code;
                    table.Name = Truncate(type.Code + " " + amount.ToString(CultureInfo.InvariantCulture), MaxUdtName);
                    table.UserFields.Fields.Item(Db.F_Trans_Session).Value = session.Code;
                    table.UserFields.Fields.Item(Db.F_Trans_Time).Value = DateTime.Now.ToString("HH:mm:ss", CultureInfo.InvariantCulture);
                    table.UserFields.Fields.Item(Db.F_Trans_Dir).Value = EnumCodes.ToCode(direction);
                    table.UserFields.Fields.Item(Db.F_Trans_TType).Value = type.Code;
                    table.UserFields.Fields.Item(Db.F_Trans_Amount).Value = amount;
                    table.UserFields.Fields.Item(Db.F_Trans_CardCode).Value = cardCode;
                    table.UserFields.Fields.Item(Db.F_Trans_Descript).Value = description;
                    table.UserFields.Fields.Item(Db.F_Trans_JE).Value = jeDocEntry;

                    int rc = table.Add();
                    DiCompany.ThrowIfError(rc, "Enregistrement de la transaction de caisse");
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
                }

                UpdateSessionFields(session.Code, theoAfter, null, null, false);

                company.EndTransaction(BoWfTransOpt.wf_Commit);
            }
            catch
            {
                if (company.InTransaction)
                    company.EndTransaction(BoWfTransOpt.wf_RollBack);
                throw;
            }
        }

        /// <summary>
        /// Clôture la session : enregistre le solde compté, calcule l'écart
        /// par rapport au solde théorique recalculé et, s'il est non nul, le
        /// comptabilise sur le compte d'écart dédié.
        /// </summary>
        public static SessionRow CloseSession(string sessionCode, double countedBalance)
        {
            SessionRow session = LoadSession(sessionCode);
            if (session == null)
                throw new InvalidOperationException("Session de caisse introuvable.");
            if (session.Status != SessionStatus.Open)
                throw new InvalidOperationException("Cette session est déjà clôturée.");

            countedBalance = Round(countedBalance);
            if (countedBalance < 0)
                throw new ArgumentException("Le solde compté ne peut pas être négatif.");

            double theo = ComputeTheoBalance(session);
            double diff = Round(countedBalance - theo);

            Company company = DiCompany.Instance;
            company.StartTransaction();
            try
            {
                if (diff != 0)
                {
                    CashSettings settings = GetSettings();
                    // Excédent (compté > théorique) : Débit Caisse / Crédit Écarts.
                    // Manquant (compté < théorique) : Débit Écarts / Crédit Caisse.
                    if (diff > 0)
                        PostJournalEntry(settings.CashAccount, settings.DiffAccount, diff, "Caisse ecart cloture (excedent)", session.Code);
                    else
                        PostJournalEntry(settings.DiffAccount, settings.CashAccount, -diff, "Caisse ecart cloture (manquant)", session.Code);
                }

                UpdateSessionFields(session.Code, theo, countedBalance, diff, true);

                company.EndTransaction(BoWfTransOpt.wf_Commit);
            }
            catch
            {
                if (company.InTransaction)
                    company.EndTransaction(BoWfTransOpt.wf_RollBack);
                throw;
            }

            return LoadSession(session.Code);
        }

        private static void UpdateSessionFields(string sessionCode, double theoBalance, double? countedBalance, double? diff, bool close)
        {
            UserTable table = DiCompany.Instance.UserTables.Item(Db.SessionTable);
            try
            {
                if (!table.GetByKey(sessionCode))
                    throw new InvalidOperationException("Session de caisse " + sessionCode + " introuvable.");

                table.UserFields.Fields.Item(Db.F_Session_TheoBal).Value = theoBalance;
                if (countedBalance.HasValue)
                    table.UserFields.Fields.Item(Db.F_Session_CountBal).Value = countedBalance.Value;
                if (diff.HasValue)
                    table.UserFields.Fields.Item(Db.F_Session_Diff).Value = diff.Value;
                if (close)
                {
                    table.UserFields.Fields.Item(Db.F_Session_Status).Value = "C";
                    table.UserFields.Fields.Item(Db.F_Session_ClosedAt).Value = DateTime.Today;
                }

                int rc = table.Update();
                DiCompany.ThrowIfError(rc, close ? "Clôture de la session de caisse" : "Mise à jour du solde théorique de la session");
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
            }
        }

        // ---------- Transactions d'une session ----------

        public static List<TransactionRow> GetTransactions(string sessionCode)
        {
            var result = new List<TransactionRow>();
            Recordset rs = NewRecordset();
            try
            {
                // Tri sur le Code (horodatage complet) et non sur l'heure seule,
                // pour rester correct si un quart du soir passe minuit.
                string sql = "SELECT T0.\"Code\", T0.\"" + Db.F_Trans_Session + "\", T0.\"" + Db.F_Trans_Time + "\", T0.\"" +
                             Db.F_Trans_Dir + "\", T0.\"" + Db.F_Trans_TType + "\", COALESCE(T1.\"Name\", T0.\"" + Db.F_Trans_TType + "\"), T0.\"" +
                             Db.F_Trans_Amount + "\", T0.\"" + Db.F_Trans_CardCode + "\", T0.\"" + Db.F_Trans_Descript + "\", T0.\"" + Db.F_Trans_JE + "\" " +
                             "FROM \"@" + Db.TransTable + "\" T0 " +
                             "LEFT JOIN \"@" + Db.TTypeTable + "\" T1 ON T1.\"Code\" = T0.\"" + Db.F_Trans_TType + "\" " +
                             "WHERE T0.\"" + Db.F_Trans_Session + "\" = '" + Sql(sessionCode) + "' " +
                             "ORDER BY T0.\"Code\"";
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
                        TransTypeName = Convert.ToString(rs.Fields.Item(5).Value),
                        Amount = Convert.ToDouble(rs.Fields.Item(6).Value),
                        CardCode = Convert.ToString(rs.Fields.Item(7).Value),
                        Description = Convert.ToString(rs.Fields.Item(8).Value),
                        JeDocEntry = Convert.ToString(rs.Fields.Item(9).Value)
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
            Recordset rs = NewRecordset();
            try
            {
                string sql = "SELECT \"Code\" FROM \"@" + Db.SessionTable + "\" " +
                             "WHERE \"" + Db.F_Session_Date + "\" >= '" + from.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + "' " +
                             "AND \"" + Db.F_Session_Date + "\" <= '" + to.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + "'";
                if (shiftFilter.HasValue)
                    sql += " AND \"" + Db.F_Session_Shift + "\" = '" + EnumCodes.ToCode(shiftFilter.Value) + "'";
                // Code = horodatage d'ouverture : ordre chronologique inverse
                sql += " ORDER BY \"Code\" DESC";

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

        private static string PostJournalEntry(string debitAccount, string creditAccount, double amount, string memo, string sessionCode)
        {
            JournalEntries je = (JournalEntries)DiCompany.Instance.GetBusinessObject(BoObjectTypes.oJournalEntries);
            try
            {
                memo = Truncate(memo, MaxJeMemo);
                je.Memo = memo;
                je.Reference = sessionCode;
                je.ReferenceDate = DateTime.Today;
                je.TaxDate = DateTime.Today;
                je.DueDate = DateTime.Today;

                je.Lines.AccountCode = debitAccount;
                je.Lines.Debit = amount;
                je.Lines.Credit = 0;
                je.Lines.LineMemo = memo;
                je.Lines.Add();

                je.Lines.AccountCode = creditAccount;
                je.Lines.Debit = 0;
                je.Lines.Credit = amount;
                je.Lines.LineMemo = memo;

                int rc = je.Add();
                DiCompany.ThrowIfError(rc, "Comptabilisation de l'écriture de caisse");

                return DiCompany.Instance.GetNewObjectKey();
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(je);
            }
        }

        // ---------- Utilitaires ----------

        private static double GetAccountBalance(string acctCode)
        {
            Recordset rs = NewRecordset();
            try
            {
                rs.DoQuery("SELECT \"CurrTotal\" FROM \"OACT\" WHERE \"AcctCode\" = '" + Sql(acctCode) + "'");
                if (rs.EoF)
                    throw new InvalidOperationException("Le compte G/L Caisse « " + acctCode + " » n'existe pas.");
                return Convert.ToDouble(rs.Fields.Item(0).Value);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }

        private static bool BusinessPartnerExists(string cardCode)
        {
            Recordset rs = NewRecordset();
            try
            {
                rs.DoQuery("SELECT \"CardCode\" FROM \"OCRD\" WHERE \"CardCode\" = '" + Sql(cardCode) + "'");
                return !rs.EoF;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }

        private static int? _sumDecimals;
        private static NumberFormatInfo _amountFormat;

        /// <summary>
        /// Lit une fois le format des montants de la société (OADM) : nombre de
        /// décimales, séparateur décimal et séparateur de milliers.
        /// </summary>
        private static void EnsureAmountFormat()
        {
            if (_sumDecimals.HasValue)
                return;

            _sumDecimals = 2;
            _amountFormat = (NumberFormatInfo)CultureInfo.CurrentCulture.NumberFormat.Clone();
            Recordset rs = NewRecordset();
            try
            {
                rs.DoQuery("SELECT \"SumDec\", \"DecSep\", \"ThousSep\" FROM \"OADM\"");
                if (!rs.EoF)
                {
                    _sumDecimals = Math.Max(0, Math.Min(6, Convert.ToInt32(rs.Fields.Item(0).Value)));
                    string decSep = Convert.ToString(rs.Fields.Item(1).Value);
                    string thousSep = Convert.ToString(rs.Fields.Item(2).Value);
                    if (!string.IsNullOrEmpty(decSep))
                        _amountFormat.NumberDecimalSeparator = decSep;
                    _amountFormat.NumberGroupSeparator = thousSep ?? "";
                }
            }
            catch
            {
                // on garde 2 décimales et le format Windows par défaut
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }

        /// <summary>Arrondit au nombre de décimales "Montants" de la société (OADM.SumDec).</summary>
        private static double Round(double value)
        {
            EnsureAmountFormat();
            return Math.Round(value, _sumDecimals.Value, MidpointRounding.AwayFromZero);
        }

        /// <summary>Montant mis en forme selon le paramétrage de la société (ex. "139 500").</summary>
        public static string FormatAmount(double value)
        {
            EnsureAmountFormat();
            return Round(value).ToString("N" + _sumDecimals.Value, _amountFormat);
        }

        /// <summary>
        /// Lit un montant saisi par l'utilisateur : accepte le format de la société
        /// (séparateurs de milliers et décimal), les espaces, et le point ou la virgule
        /// comme séparateur décimal.
        /// </summary>
        public static bool TryParseAmount(string text, out double value)
        {
            EnsureAmountFormat();
            value = 0;
            if (string.IsNullOrWhiteSpace(text))
                return false;

            string s = text.Trim().Replace(" ", "").Replace(" ", "").Replace(" ", "");
            string groupSep = _amountFormat.NumberGroupSeparator;
            if (!string.IsNullOrEmpty(groupSep) && groupSep != _amountFormat.NumberDecimalSeparator)
                s = s.Replace(groupSep, "");
            s = s.Replace(_amountFormat.NumberDecimalSeparator, ".").Replace(",", ".");

            return double.TryParse(s, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out value);
        }

        private static Recordset NewRecordset()
        {
            return (Recordset)DiCompany.Instance.GetBusinessObject(BoObjectTypes.BoRecordset);
        }

        private static string Sql(string value)
        {
            return (value ?? "").Replace("'", "''");
        }

        private static string Truncate(string value, int max)
        {
            if (value == null)
                return "";
            return value.Length > max ? value.Substring(0, max) : value;
        }
    }
}
