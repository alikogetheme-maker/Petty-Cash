using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
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
        public string UserCode;
        public string CashBox;
        public double ExpOpen;
        public double OpenDiff;
        public double OpenBal;
        public double TheoBal;
        public double CountBal;
        public double Diff;
        public SessionStatus Status;
        public DateTime? ClosedAt;
        public string ClosedBy;
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

    internal class CashBox
    {
        public string Code;
        public string Name;
        public string CashAccount;
        public string DiffAccount;      // compte d'écarts effectif (propre à la caisse ou par défaut)
        public bool Active;
        public string OpenSessionCode;  // null si la caisse est libre
        public string OpenedBy;         // nom de l'utilisateur qui l'occupe

        public bool IsFree => Active && OpenSessionCode == null;
    }

    internal class Denomination
    {
        public string Code;
        public string Name;
        public double Value;
        public string Kind;
    }

    internal class CountLine
    {
        public string DenomCode;
        public string Label;
        public double Value;
        public int Qty;
        public double Amount => Value * Qty;
    }

    internal class SapUser
    {
        public string Code;
        public string Name;
        public bool IsSuperUser;
    }

    /// <summary>
    /// Toute la logique métier de la caisse. C'est le fichier qui fait
    /// référence pour comprendre le comportement fonctionnel de l'add-on.
    ///
    /// Règles :
    /// - plusieurs caisses (table @PC_CASHBOX), chacune avec son compte G/L ;
    /// - une session = une caisse + un utilisateur SAP + un quart ;
    /// - une caisse n'a qu'une session ouverte, un utilisateur n'a qu'une
    ///   session ouverte (verrous @PC_LOCK protégés par l'index unique SAP) ;
    /// - billetage obligatoire à l'ouverture et à la clôture ; l'écart
    ///   d'ouverture (billetage − solde repris) et l'écart de clôture
    ///   (billetage − théorique) sont comptabilisés sur le compte d'écarts ;
    /// - seul l'ouvreur saisit des opérations ; l'ouvreur ou un
    ///   superutilisateur SAP peut clôturer ;
    /// - solde d'ouverture attendu = compté de la dernière session de la
    ///   caisse (ou, la première fois, solde du compte G/L de la caisse) ;
    /// - solde théorique = ouverture + recettes − dépenses, recalculé en base ;
    /// - une dépense ne peut pas rendre le solde théorique négatif ;
    /// - toute opération (écriture + lignes + session + verrous) est faite
    ///   dans une même transaction DI API (tout ou rien).
    /// </summary>
    internal static class CashSessionService
    {
        // Longueurs maximales des champs SAP / UDF alimentés
        private const int MaxJeMemo = 50;        // OJDT.Memo, JDT1.LineMemo
        private const int MaxCashier = 50;
        private const int MaxDescription = 100;
        private const int DuplicateKeyError = -2035;

        // ---------- Utilisateur connecté ----------

        private static SapUser _currentUser;

        /// <summary>Utilisateur SAP connecté (celui du client SAP B1 qui a lancé l'add-on).</summary>
        public static SapUser GetCurrentUser()
        {
            if (_currentUser != null)
                return _currentUser;

            string code = DiCompany.Instance.UserName;
            var user = new SapUser { Code = code, Name = code };
            Recordset rs = NewRecordset();
            try
            {
                rs.DoQuery("SELECT \"U_NAME\", \"SUPERUSER\" FROM \"OUSR\" WHERE \"USER_CODE\" = '" + Sql(code) + "'");
                if (!rs.EoF)
                {
                    string name = Convert.ToString(rs.Fields.Item(0).Value).Trim();
                    if (name.Length > 0)
                        user.Name = name;
                    user.IsSuperUser = Convert.ToString(rs.Fields.Item(1).Value) == "Y";
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
            _currentUser = user;
            return user;
        }

        private static bool SameUser(string a, string b)
        {
            return string.Equals((a ?? "").Trim(), (b ?? "").Trim(), StringComparison.OrdinalIgnoreCase);
        }

        // ---------- Paramétrage ----------

        public static CashSettings GetSettings()
        {
            UserTable table = DiCompany.Instance.UserTables.Item(Db.SettingsTable);
            try
            {
                if (!table.GetByKey(Db.SettingsCode))
                    throw new InvalidOperationException("Paramétrage caisse introuvable (table @" + Db.SettingsTable + ").");

                return new CashSettings
                {
                    CashAccount = Convert.ToString(table.UserFields.Fields.Item(Db.F_CashAcct).Value).Trim(),
                    DiffAccount = Convert.ToString(table.UserFields.Fields.Item(Db.F_DiffAcct).Value).Trim()
                };
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

        // ---------- Caisses ----------

        /// <summary>Caisses avec, pour chacune, la session qui l'occupe éventuellement.</summary>
        public static List<CashBox> GetCashBoxes(bool activeOnly)
        {
            string defaultDiff = GetSettings().DiffAccount;
            var result = new List<CashBox>();
            Recordset rs = NewRecordset();
            try
            {
                string sql = "SELECT B.\"Code\", B.\"Name\", B.\"" + Db.F_Box_CashAcct + "\", B.\"" + Db.F_Box_DiffAcct + "\", B.\"" + Db.F_Box_Active + "\", " +
                             "S.\"Code\", S.\"" + Db.F_Session_Cashier + "\" " +
                             "FROM \"@" + Db.CashBoxTable + "\" B " +
                             "LEFT JOIN \"@" + Db.SessionTable + "\" S ON S.\"" + Db.F_Session_CashBox + "\" = B.\"Code\" AND S.\"" + Db.F_Session_Status + "\" = 'O' " +
                             (activeOnly ? "WHERE B.\"" + Db.F_Box_Active + "\" = 'Y' " : "") +
                             "ORDER BY B.\"Name\", S.\"Code\" DESC";
                rs.DoQuery(sql);
                while (!rs.EoF)
                {
                    string code = Convert.ToString(rs.Fields.Item(0).Value);
                    if (!result.Any(b => b.Code == code))
                    {
                        string boxDiff = Convert.ToString(rs.Fields.Item(3).Value).Trim();
                        string sessCode = Convert.ToString(rs.Fields.Item(5).Value);
                        result.Add(new CashBox
                        {
                            Code = code,
                            Name = Convert.ToString(rs.Fields.Item(1).Value),
                            CashAccount = Convert.ToString(rs.Fields.Item(2).Value).Trim(),
                            DiffAccount = boxDiff.Length > 0 ? boxDiff : defaultDiff,
                            Active = Convert.ToString(rs.Fields.Item(4).Value) == "Y",
                            OpenSessionCode = string.IsNullOrEmpty(sessCode) ? null : sessCode,
                            OpenedBy = Convert.ToString(rs.Fields.Item(6).Value)
                        });
                    }
                    rs.MoveNext();
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
            return result;
        }

        public static CashBox GetCashBox(string code)
        {
            return GetCashBoxes(false).FirstOrDefault(b => b.Code == code);
        }

        private static CashBox RequireCashBoxAccounts(string code)
        {
            CashBox box = GetCashBox(code);
            if (box == null)
                throw new InvalidOperationException("Caisse « " + code + " » introuvable (table @" + Db.CashBoxTable + ").");
            if (string.IsNullOrEmpty(box.CashAccount))
                throw new InvalidOperationException("Aucun compte G/L n'est paramétré pour la caisse « " + box.Name + " ».");
            if (string.IsNullOrEmpty(box.DiffAccount))
                throw new InvalidOperationException("Renseignez le compte d'écarts de caisse (U_DiffAcct) dans @" + Db.SettingsTable + " ou sur la caisse « " + box.Name + " ».");
            return box;
        }

        // ---------- Coupures et billetage ----------

        public static List<Denomination> GetDenominations()
        {
            var result = new List<Denomination>();
            Recordset rs = NewRecordset();
            try
            {
                rs.DoQuery("SELECT \"Code\", \"Name\", \"" + Db.F_Denom_Value + "\", \"" + Db.F_Denom_Kind + "\" FROM \"@" + Db.DenomTable + "\" " +
                           "WHERE \"" + Db.F_Denom_Active + "\" = 'Y' ORDER BY \"" + Db.F_Denom_Kind + "\", \"" + Db.F_Denom_Value + "\" DESC");
                while (!rs.EoF)
                {
                    result.Add(new Denomination
                    {
                        Code = Convert.ToString(rs.Fields.Item(0).Value),
                        Name = Convert.ToString(rs.Fields.Item(1).Value),
                        Value = Convert.ToDouble(rs.Fields.Item(2).Value),
                        Kind = Convert.ToString(rs.Fields.Item(3).Value)
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

        /// <summary>Lignes de billetage d'une session pour une phase ('O' ouverture, 'C' clôture).</summary>
        public static List<CountLine> GetCountLines(string sessionCode, string phase)
        {
            var result = new List<CountLine>();
            Recordset rs = NewRecordset();
            try
            {
                rs.DoQuery("SELECT C.\"" + Db.F_Count_Denom + "\", COALESCE(D.\"Name\", C.\"" + Db.F_Count_Denom + "\"), C.\"" + Db.F_Count_Value + "\", C.\"" + Db.F_Count_Qty + "\" " +
                           "FROM \"@" + Db.CountTable + "\" C LEFT JOIN \"@" + Db.DenomTable + "\" D ON D.\"Code\" = C.\"" + Db.F_Count_Denom + "\" " +
                           "WHERE C.\"" + Db.F_Count_Session + "\" = '" + Sql(sessionCode) + "' AND C.\"" + Db.F_Count_Phase + "\" = '" + Sql(phase) + "' " +
                           "ORDER BY C.\"" + Db.F_Count_Value + "\" DESC, C.\"" + Db.F_Count_Denom + "\"");
                while (!rs.EoF)
                {
                    result.Add(new CountLine
                    {
                        DenomCode = Convert.ToString(rs.Fields.Item(0).Value),
                        Label = Convert.ToString(rs.Fields.Item(1).Value),
                        Value = Convert.ToDouble(rs.Fields.Item(2).Value),
                        Qty = Convert.ToInt32(rs.Fields.Item(3).Value)
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

        private static double CountTotal(List<CountLine> lines)
        {
            if (lines == null)
                throw new ArgumentException("Le billetage est obligatoire.");
            if (lines.Any(l => l.Qty < 0))
                throw new ArgumentException("Une quantité de billetage ne peut pas être négative.");
            return Round(lines.Sum(l => l.Amount));
        }

        private static void SaveCountLines(string sessionCode, string phase, List<CountLine> lines)
        {
            foreach (var line in lines.Where(l => l.Qty > 0))
            {
                UserTable table = DiCompany.Instance.UserTables.Item(Db.CountTable);
                try
                {
                    string code = sessionCode + phase + line.DenomCode;
                    table.Code = code;
                    table.Name = code;
                    table.UserFields.Fields.Item(Db.F_Count_Session).Value = sessionCode;
                    table.UserFields.Fields.Item(Db.F_Count_Phase).Value = phase;
                    table.UserFields.Fields.Item(Db.F_Count_Denom).Value = line.DenomCode;
                    table.UserFields.Fields.Item(Db.F_Count_Value).Value = line.Value;
                    table.UserFields.Fields.Item(Db.F_Count_Qty).Value = line.Qty;
                    table.UserFields.Fields.Item(Db.F_Count_Amount).Value = Round(line.Amount);
                    DiCompany.ThrowIfError(table.Add(), "Enregistrement du billetage");
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
                }
            }
        }

        // ---------- Verrous ----------

        private static string CashBoxLockKey(string cashBox) { return "CB:" + cashBox; }
        private static string UserLockKey(string userCode) { return "US:" + userCode.ToUpperInvariant(); }

        /// <summary>Pose un verrou ; false si la clé existe déjà (index unique SAP).</summary>
        private static bool TryAddLock(string key, string sessionCode)
        {
            Company company = DiCompany.Instance;
            UserTable table = company.UserTables.Item(Db.LockTable);
            try
            {
                table.Code = key;
                table.Name = key;
                table.UserFields.Fields.Item(Db.F_Lock_Session).Value = sessionCode;
                int rc = table.Add();
                if (rc == 0)
                    return true;
                company.GetLastError(out int errCode, out string errMsg);
                if (errCode == DuplicateKeyError)
                    return false;
                throw new InvalidOperationException("Pose du verrou " + key + " : [" + errCode + "] " + errMsg);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
            }
        }

        private static void RemoveLock(string key)
        {
            UserTable table = DiCompany.Instance.UserTables.Item(Db.LockTable);
            try
            {
                if (table.GetByKey(key))
                    DiCompany.ThrowIfError(table.Remove(), "Suppression du verrou " + key);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
            }
        }

        private static List<string> GetLockKeys(string sessionCode)
        {
            var keys = new List<string>();
            Recordset rs = NewRecordset();
            try
            {
                rs.DoQuery("SELECT \"Code\" FROM \"@" + Db.LockTable + "\" WHERE \"" + Db.F_Lock_Session + "\" = '" + Sql(sessionCode) + "'");
                while (!rs.EoF)
                {
                    keys.Add(Convert.ToString(rs.Fields.Item(0).Value));
                    rs.MoveNext();
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
            return keys;
        }

        /// <summary>
        /// Supprime un verrou dont la session n'est plus ouverte (session
        /// clôturée ou supprimée hors de l'add-on) pour ne pas bloquer une caisse.
        /// </summary>
        private static void ReleaseStaleLock(string key)
        {
            string sessionCode = Scalar("SELECT \"" + Db.F_Lock_Session + "\" FROM \"@" + Db.LockTable + "\" WHERE \"Code\" = '" + Sql(key) + "'");
            if (sessionCode == null)
                return;
            SessionRow session = LoadSession(sessionCode);
            if (session == null || session.Status != SessionStatus.Open)
                RemoveLock(key);
        }

        // ---------- Sessions ----------

        /// <summary>Session ouverte de l'utilisateur (null s'il n'en a pas).</summary>
        public static SessionRow GetOpenSessionForUser(string userCode)
        {
            string code = Scalar("SELECT \"Code\" FROM \"@" + Db.SessionTable + "\" WHERE \"" + Db.F_Session_Status + "\" = 'O' " +
                                 "AND UPPER(\"" + Db.F_Session_User + "\") = '" + Sql(userCode.ToUpperInvariant()) + "' ORDER BY \"Code\" DESC");
            return code == null ? null : LoadSession(code);
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

            Fields f = table.UserFields.Fields;
            return new SessionRow
            {
                Code = table.Code,
                CashDate = Convert.ToDateTime(f.Item(Db.F_Session_Date).Value),
                Shift = EnumCodes.ShiftFromCode(Convert.ToString(f.Item(Db.F_Session_Shift).Value)),
                Cashier = Convert.ToString(f.Item(Db.F_Session_Cashier).Value),
                UserCode = Convert.ToString(f.Item(Db.F_Session_User).Value),
                CashBox = Convert.ToString(f.Item(Db.F_Session_CashBox).Value),
                ExpOpen = Convert.ToDouble(f.Item(Db.F_Session_ExpOpen).Value),
                OpenDiff = Convert.ToDouble(f.Item(Db.F_Session_OpenDiff).Value),
                OpenBal = Convert.ToDouble(f.Item(Db.F_Session_OpenBal).Value),
                TheoBal = Convert.ToDouble(f.Item(Db.F_Session_TheoBal).Value),
                CountBal = Convert.ToDouble(f.Item(Db.F_Session_CountBal).Value),
                Diff = Convert.ToDouble(f.Item(Db.F_Session_Diff).Value),
                Status = EnumCodes.StatusFromCode(Convert.ToString(f.Item(Db.F_Session_Status).Value)),
                ClosedAt = closedAt,
                ClosedBy = Convert.ToString(f.Item(Db.F_Session_ClosedBy).Value)
            };
        }

        /// <summary>
        /// Solde attendu à la prochaine ouverture de la caisse : solde compté de
        /// sa dernière session clôturée, ou, s'il n'y en a encore aucune, solde
        /// actuel de son compte G/L (le fonds initial doit y être comptabilisé).
        /// </summary>
        public static double GetNextOpeningBalance(string cashBoxCode)
        {
            string last = Scalar("SELECT TOP 1 \"" + Db.F_Session_CountBal + "\" FROM \"@" + Db.SessionTable + "\" " +
                                 "WHERE \"" + Db.F_Session_Status + "\" = 'C' AND \"" + Db.F_Session_CashBox + "\" = '" + Sql(cashBoxCode) + "' " +
                                 "ORDER BY \"Code\" DESC");
            if (last != null)
                return Round(Convert.ToDouble(last, CultureInfo.InvariantCulture));

            CashBox box = GetCashBox(cashBoxCode);
            if (box == null)
                throw new InvalidOperationException("Caisse « " + cashBoxCode + " » introuvable.");
            return Round(GetAccountBalance(box.CashAccount));
        }

        /// <summary>
        /// Ouvre une session sur une caisse libre pour l'utilisateur connecté.
        /// Le solde d'ouverture est le billetage saisi ; s'il diffère du solde
        /// attendu, l'écart d'ouverture est comptabilisé.
        /// </summary>
        public static SessionRow OpenSession(string cashBoxCode, Shift shift, List<CountLine> openingCount)
        {
            SapUser user = GetCurrentUser();
            CashBox box = RequireCashBoxAccounts(cashBoxCode);
            if (!box.Active)
                throw new InvalidOperationException("La caisse « " + box.Name + " » est désactivée.");

            SessionRow mine = GetOpenSessionForUser(user.Code);
            if (mine != null)
                throw new InvalidOperationException("Vous avez déjà une session ouverte sur la caisse « " + CashBoxName(mine.CashBox) +
                                                    " ». Clôturez-la avant d'en ouvrir une autre.");
            if (box.OpenSessionCode != null)
                throw new InvalidOperationException("La caisse « " + box.Name + " » est déjà ouverte par " + box.OpenedBy + ".");

            double counted = CountTotal(openingCount);
            double expected = GetNextOpeningBalance(box.Code);
            double openDiff = Round(counted - expected);

            ReleaseStaleLock(CashBoxLockKey(box.Code));
            ReleaseStaleLock(UserLockKey(user.Code));

            string code = DateTime.Now.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture);
            Company company = DiCompany.Instance;
            company.StartTransaction();
            try
            {
                if (!TryAddLock(CashBoxLockKey(box.Code), code))
                    throw new InvalidOperationException("La caisse « " + box.Name + " » vient d'être ouverte par un autre utilisateur. Choisissez une autre caisse.");
                if (!TryAddLock(UserLockKey(user.Code), code))
                    throw new InvalidOperationException("Vous avez déjà une session ouverte sur une autre caisse. Clôturez-la avant d'en ouvrir une autre.");

                if (openDiff > 0)
                    PostJournalEntry(box.CashAccount, box.DiffAccount, openDiff, "Caisse ecart ouverture (excedent) " + box.Code, code);
                else if (openDiff < 0)
                    PostJournalEntry(box.DiffAccount, box.CashAccount, -openDiff, "Caisse ecart ouverture (manquant) " + box.Code, code);

                UserTable table = company.UserTables.Item(Db.SessionTable);
                try
                {
                    Fields f = table.UserFields.Fields;
                    table.Code = code;
                    table.Name = code;   // index unique SAP sur Name
                    f.Item(Db.F_Session_Date).Value = DateTime.Today;
                    f.Item(Db.F_Session_Shift).Value = EnumCodes.ToCode(shift);
                    f.Item(Db.F_Session_Cashier).Value = Truncate(user.Name, MaxCashier);
                    f.Item(Db.F_Session_User).Value = user.Code;
                    f.Item(Db.F_Session_CashBox).Value = box.Code;
                    f.Item(Db.F_Session_ExpOpen).Value = expected;
                    f.Item(Db.F_Session_OpenDiff).Value = openDiff;
                    f.Item(Db.F_Session_OpenBal).Value = counted;
                    f.Item(Db.F_Session_TheoBal).Value = counted;
                    f.Item(Db.F_Session_CountBal).Value = 0;
                    f.Item(Db.F_Session_Diff).Value = 0;
                    f.Item(Db.F_Session_Status).Value = "O";
                    DiCompany.ThrowIfError(table.Add(), "Ouverture de la session de caisse");
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
                }

                SaveCountLines(code, "O", openingCount);

                company.EndTransaction(BoWfTransOpt.wf_Commit);
            }
            catch
            {
                if (company.InTransaction)
                    company.EndTransaction(BoWfTransOpt.wf_RollBack);
                throw;
            }

            return LoadSession(code);
        }

        public static string CashBoxName(string cashBoxCode)
        {
            return Scalar("SELECT \"Name\" FROM \"@" + Db.CashBoxTable + "\" WHERE \"Code\" = '" + Sql(cashBoxCode) + "'") ?? cashBoxCode;
        }

        /// <summary>
        /// Solde théorique recalculé depuis la base : ouverture + recettes - dépenses.
        /// </summary>
        private static double ComputeTheoBalance(SessionRow session)
        {
            string sum = Scalar("SELECT COALESCE(SUM(CASE WHEN \"" + Db.F_Trans_Dir + "\" = 'R' THEN \"" + Db.F_Trans_Amount + "\" " +
                                "ELSE -\"" + Db.F_Trans_Amount + "\" END), 0) " +
                                "FROM \"@" + Db.TransTable + "\" WHERE \"" + Db.F_Trans_Session + "\" = '" + Sql(session.Code) + "'");
            return Round(session.OpenBal + Convert.ToDouble(sum, CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Ajoute une transaction à une session ouverte de l'utilisateur
        /// connecté : écriture au journal (caisse vs compte du type), ligne de
        /// caisse et solde théorique, dans une même transaction DI API.
        /// </summary>
        public static void AddTransaction(string sessionCode, Direction direction, TransactionType type, double amount, string cardCode, string description)
        {
            SessionRow session = LoadSession(sessionCode);
            if (session == null)
                throw new InvalidOperationException("Session de caisse introuvable.");
            if (session.Status != SessionStatus.Open)
                throw new InvalidOperationException("Cette session est clôturée, impossible d'ajouter une transaction.");

            SapUser user = GetCurrentUser();
            bool legacyNoOwner = string.IsNullOrEmpty(session.UserCode) && user.IsSuperUser;
            if (!SameUser(session.UserCode, user.Code) && !legacyNoOwner)
                throw new InvalidOperationException("Seul l'utilisateur qui a ouvert cette session (" + session.Cashier + ") peut y saisir des opérations.");

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

            CashBox box = RequireCashBoxAccounts(session.CashBox);

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
                    direction == Direction.Recette ? box.CashAccount : type.GlAccount,
                    direction == Direction.Recette ? type.GlAccount : box.CashAccount,
                    amount, memo, session.Code);

                string code = DateTime.Now.ToString("yyyyMMddHHmmssfff", CultureInfo.InvariantCulture);
                UserTable table = company.UserTables.Item(Db.TransTable);
                try
                {
                    table.Code = code;
                    // Index unique SAP sur Name : "type + montant" provoquait l'erreur -2035
                    // dès qu'une même opération du même montant était ressaisie.
                    table.Name = code;
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

                UpdateSessionFields(session.Code, theoAfter, null, null, null);

                company.EndTransaction(BoWfTransOpt.wf_Commit);
            }
            catch
            {
                if (company.InTransaction)
                    company.EndTransaction(BoWfTransOpt.wf_RollBack);
                throw;
            }
        }

        /// <summary>Indique si l'utilisateur connecté peut clôturer cette session.</summary>
        public static bool CanClose(SessionRow session)
        {
            SapUser user = GetCurrentUser();
            return session != null && session.Status == SessionStatus.Open &&
                   (SameUser(session.UserCode, user.Code) || user.IsSuperUser);
        }

        /// <summary>
        /// Clôture la session : le billetage donne le solde compté, l'écart avec
        /// le théorique recalculé est comptabilisé, la caisse est libérée.
        /// Réservé à l'ouvreur ou à un superutilisateur SAP.
        /// </summary>
        public static SessionRow CloseSession(string sessionCode, List<CountLine> closingCount)
        {
            SessionRow session = LoadSession(sessionCode);
            if (session == null)
                throw new InvalidOperationException("Session de caisse introuvable.");
            if (session.Status != SessionStatus.Open)
                throw new InvalidOperationException("Cette session est déjà clôturée.");
            if (!CanClose(session))
                throw new InvalidOperationException("Seul l'utilisateur qui a ouvert cette session (" + session.Cashier + ") ou un superutilisateur peut la clôturer.");

            double countedBalance = CountTotal(closingCount);
            CashBox box = RequireCashBoxAccounts(session.CashBox);
            double theo = ComputeTheoBalance(session);
            double diff = Round(countedBalance - theo);

            Company company = DiCompany.Instance;
            company.StartTransaction();
            try
            {
                // Excédent (compté > théorique) : Débit Caisse / Crédit Écarts.
                // Manquant (compté < théorique) : Débit Écarts / Crédit Caisse.
                if (diff > 0)
                    PostJournalEntry(box.CashAccount, box.DiffAccount, diff, "Caisse ecart cloture (excedent) " + box.Code, session.Code);
                else if (diff < 0)
                    PostJournalEntry(box.DiffAccount, box.CashAccount, -diff, "Caisse ecart cloture (manquant) " + box.Code, session.Code);

                UpdateSessionFields(session.Code, theo, countedBalance, diff, GetCurrentUser().Code);
                SaveCountLines(session.Code, "C", closingCount);

                foreach (string key in GetLockKeys(session.Code))
                    RemoveLock(key);

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

        /// <summary>Mise à jour de la session ; closedBy non null = clôture.</summary>
        private static void UpdateSessionFields(string sessionCode, double theoBalance, double? countedBalance, double? diff, string closedBy)
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
                if (closedBy != null)
                {
                    table.UserFields.Fields.Item(Db.F_Session_Status).Value = "C";
                    table.UserFields.Fields.Item(Db.F_Session_ClosedAt).Value = DateTime.Today;
                    table.UserFields.Fields.Item(Db.F_Session_ClosedBy).Value = closedBy;
                }

                int rc = table.Update();
                DiCompany.ThrowIfError(rc, closedBy != null ? "Clôture de la session de caisse" : "Mise à jour du solde théorique de la session");
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

        // ---------- Rapports ----------

        public static List<SessionRow> GetSessions(DateTime from, DateTime to, Shift? shiftFilter, string cashBoxFilter = null)
        {
            var codes = new List<string>();
            Recordset rs = NewRecordset();
            try
            {
                string sql = "SELECT \"Code\" FROM \"@" + Db.SessionTable + "\" " +
                             "WHERE \"" + Db.F_Session_Date + "\" >= '" + from.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + "' " +
                             "AND \"" + Db.F_Session_Date + "\" <= '" + to.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + "'";
                if (shiftFilter.HasValue)
                    sql += " AND \"" + Db.F_Session_Shift + "\" = '" + EnumCodes.ToCode(shiftFilter.Value) + "'";
                if (!string.IsNullOrEmpty(cashBoxFilter))
                    sql += " AND \"" + Db.F_Session_CashBox + "\" = '" + Sql(cashBoxFilter) + "'";
                // Code = horodatage d'ouverture : ordre chronologique inverse
                sql += " ORDER BY \"Code\" DESC";

                rs.DoQuery(sql);
                while (!rs.EoF)
                {
                    codes.Add(Convert.ToString(rs.Fields.Item(0).Value));
                    rs.MoveNext();
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
            return codes.Select(LoadSession).ToList();
        }

        /// <summary>Utilisateurs ayant déjà tenu une session (filtre du rapport) : code -> nom.</summary>
        public static List<KeyValuePair<string, string>> GetSessionUsers()
        {
            var result = new List<KeyValuePair<string, string>>();
            Recordset rs = NewRecordset();
            try
            {
                rs.DoQuery("SELECT DISTINCT S.\"" + Db.F_Session_User + "\", COALESCE(U.\"U_NAME\", S.\"" + Db.F_Session_User + "\") " +
                           "FROM \"@" + Db.SessionTable + "\" S LEFT JOIN \"OUSR\" U ON U.\"USER_CODE\" = S.\"" + Db.F_Session_User + "\" " +
                           "WHERE COALESCE(S.\"" + Db.F_Session_User + "\", '') <> ''");
                while (!rs.EoF)
                {
                    result.Add(new KeyValuePair<string, string>(Convert.ToString(rs.Fields.Item(0).Value), Convert.ToString(rs.Fields.Item(1).Value)));
                    rs.MoveNext();
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
            return result.OrderBy(kv => kv.Value).ToList();
        }

        /// <summary>Filtre WHERE commun aux requêtes de rapport (alias S = sessions).</summary>
        private static string ReportWhere(DateTime from, DateTime to, string cashBox, string userCode, Shift? shift)
        {
            string where = "WHERE S.\"" + Db.F_Session_Date + "\" >= '" + from.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + "' " +
                           "AND S.\"" + Db.F_Session_Date + "\" <= '" + to.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + "' ";
            if (!string.IsNullOrEmpty(cashBox))
                where += "AND S.\"" + Db.F_Session_CashBox + "\" = '" + Sql(cashBox) + "' ";
            if (!string.IsNullOrEmpty(userCode))
                where += "AND S.\"" + Db.F_Session_User + "\" = '" + Sql(userCode) + "' ";
            if (shift.HasValue)
                where += "AND S.\"" + Db.F_Session_Shift + "\" = '" + EnumCodes.ToCode(shift.Value) + "' ";
            return where;
        }

        private static string ReportFrom()
        {
            return "FROM \"@" + Db.SessionTable + "\" S " +
                   "LEFT JOIN \"@" + Db.CashBoxTable + "\" B ON B.\"Code\" = S.\"" + Db.F_Session_CashBox + "\" " +
                   "LEFT JOIN \"OUSR\" U ON U.\"USER_CODE\" = S.\"" + Db.F_Session_User + "\" " +
                   "LEFT JOIN (SELECT \"" + Db.F_Trans_Session + "\" AS \"Sess\", " +
                   "SUM(CASE WHEN \"" + Db.F_Trans_Dir + "\" = 'R' THEN \"" + Db.F_Trans_Amount + "\" ELSE 0 END) AS \"Rec\", " +
                   "SUM(CASE WHEN \"" + Db.F_Trans_Dir + "\" = 'D' THEN \"" + Db.F_Trans_Amount + "\" ELSE 0 END) AS \"Dep\" " +
                   "FROM \"@" + Db.TransTable + "\" GROUP BY \"" + Db.F_Trans_Session + "\") T ON T.\"Sess\" = S.\"Code\" ";
        }

        /// <summary>Rapport détaillé : une ligne par session (colonne Code masquée à l'écran).</summary>
        public static string SessionsReportSql(DateTime from, DateTime to, string cashBox, string userCode, Shift? shift)
        {
            return "SELECT S.\"Code\" AS \"Code\", S.\"" + Db.F_Session_Date + "\" AS \"Date\", " +
                   "COALESCE(B.\"Name\", S.\"" + Db.F_Session_CashBox + "\") AS \"Caisse\", " +
                   "CASE S.\"" + Db.F_Session_Shift + "\" WHEN 'A' THEN 'Après-midi' WHEN 'S' THEN 'Soir' ELSE 'Matin' END AS \"Quart\", " +
                   "COALESCE(U.\"U_NAME\", S.\"" + Db.F_Session_Cashier + "\") AS \"Util\", " +
                   "S.\"" + Db.F_Session_OpenBal + "\" AS \"Ouv\", S.\"" + Db.F_Session_OpenDiff + "\" AS \"EcOuv\", " +
                   "COALESCE(T.\"Rec\", 0) AS \"Rec\", COALESCE(T.\"Dep\", 0) AS \"Dep\", " +
                   "S.\"" + Db.F_Session_TheoBal + "\" AS \"Theo\", " +
                   "CASE WHEN S.\"" + Db.F_Session_Status + "\" = 'C' THEN S.\"" + Db.F_Session_CountBal + "\" ELSE 0 END AS \"Compte\", " +
                   "CASE WHEN S.\"" + Db.F_Session_Status + "\" = 'C' THEN S.\"" + Db.F_Session_Diff + "\" ELSE 0 END AS \"Ecart\", " +
                   "CASE WHEN S.\"" + Db.F_Session_Status + "\" = 'C' THEN 'Clôturée' ELSE 'Ouverte' END AS \"Statut\" " +
                   ReportFrom() + ReportWhere(from, to, cashBox, userCode, shift) +
                   "ORDER BY S.\"Code\" DESC";
        }

        /// <summary>Synthèse par caisse ou par utilisateur sur la période.</summary>
        public static string SummaryReportSql(bool byCashBox, DateTime from, DateTime to, string cashBox, string userCode, Shift? shift)
        {
            string key = byCashBox
                ? "COALESCE(B.\"Name\", S.\"" + Db.F_Session_CashBox + "\")"
                : "COALESCE(U.\"U_NAME\", S.\"" + Db.F_Session_Cashier + "\")";
            return "SELECT " + key + " AS \"" + (byCashBox ? "Caisse" : "Util") + "\", " +
                   "COUNT(*) AS \"NbSess\", " +
                   "SUM(CASE WHEN S.\"" + Db.F_Session_Status + "\" = 'O' THEN 1 ELSE 0 END) AS \"NbOuv\", " +
                   "SUM(COALESCE(T.\"Rec\", 0)) AS \"Rec\", SUM(COALESCE(T.\"Dep\", 0)) AS \"Dep\", " +
                   "SUM(S.\"" + Db.F_Session_OpenDiff + "\") AS \"EcOuv\", " +
                   "SUM(CASE WHEN S.\"" + Db.F_Session_Status + "\" = 'C' THEN S.\"" + Db.F_Session_Diff + "\" ELSE 0 END) AS \"Ecart\" " +
                   ReportFrom() + ReportWhere(from, to, cashBox, userCode, shift) +
                   "GROUP BY " + key + " ORDER BY " + key;
        }

        public static string TransactionsReportSql(string sessionCode)
        {
            return "SELECT T0.\"" + Db.F_Trans_Time + "\" AS \"Heure\", " +
                   "CASE WHEN T0.\"" + Db.F_Trans_Dir + "\" = 'R' THEN 'Recette' ELSE 'Dépense' END AS \"Sens\", " +
                   "COALESCE(T1.\"Name\", T0.\"" + Db.F_Trans_TType + "\") AS \"Type\", T0.\"" + Db.F_Trans_Amount + "\" AS \"Montant\", " +
                   "T0.\"" + Db.F_Trans_CardCode + "\" AS \"Tiers\", T0.\"" + Db.F_Trans_Descript + "\" AS \"Descr\", T0.\"" + Db.F_Trans_JE + "\" AS \"Ecriture\" " +
                   "FROM \"@" + Db.TransTable + "\" T0 LEFT JOIN \"@" + Db.TTypeTable + "\" T1 ON T1.\"Code\" = T0.\"" + Db.F_Trans_TType + "\" " +
                   "WHERE T0.\"" + Db.F_Trans_Session + "\" = '" + Sql(sessionCode) + "' ORDER BY T0.\"Code\"";
        }

        /// <summary>Billetage d'ouverture et de clôture côte à côte, par coupure.</summary>
        public static string CountReportSql(string sessionCode)
        {
            string c = "\"@" + Db.CountTable + "\"";
            return "SELECT COALESCE(D.\"Name\", X.\"Denom\") AS \"Coupure\", " +
                   "SUM(CASE WHEN X.\"Phase\" = 'O' THEN X.\"Qty\" ELSE 0 END) AS \"QtOuv\", " +
                   "SUM(CASE WHEN X.\"Phase\" = 'C' THEN X.\"Qty\" ELSE 0 END) AS \"QtClo\", " +
                   "SUM(CASE WHEN X.\"Phase\" = 'C' THEN X.\"Amt\" ELSE 0 END) AS \"MtClo\" " +
                   "FROM (SELECT \"" + Db.F_Count_Denom + "\" AS \"Denom\", \"" + Db.F_Count_Phase + "\" AS \"Phase\", \"" + Db.F_Count_Qty + "\" AS \"Qty\", " +
                   "\"" + Db.F_Count_Amount + "\" AS \"Amt\", \"" + Db.F_Count_Value + "\" AS \"Val\" FROM " + c +
                   " WHERE \"" + Db.F_Count_Session + "\" = '" + Sql(sessionCode) + "') X " +
                   "LEFT JOIN \"@" + Db.DenomTable + "\" D ON D.\"Code\" = X.\"Denom\" " +
                   "GROUP BY COALESCE(D.\"Name\", X.\"Denom\"), X.\"Val\" ORDER BY X.\"Val\" DESC, COALESCE(D.\"Name\", X.\"Denom\")";
        }

        // ---------- Migration mono-caisse -> multi-caisses ----------

        /// <summary>
        /// Idempotent. Crée la caisse "CAISSE1" à partir de l'ancien paramétrage
        /// mono-caisse, y rattache les sessions existantes et pose les verrous
        /// des sessions encore ouvertes.
        /// </summary>
        public static void MigrateToMultiCashBox()
        {
            Company company = DiCompany.Instance;

            if (Scalar("SELECT TOP 1 \"Code\" FROM \"@" + Db.CashBoxTable + "\"") == null)
            {
                string legacyAcct = GetSettings().CashAccount;
                if (string.IsNullOrEmpty(legacyAcct))
                    return;   // installation neuve : les caisses sont à créer par l'administrateur

                UserTable box = company.UserTables.Item(Db.CashBoxTable);
                try
                {
                    box.Code = "CAISSE1";
                    box.Name = "Caisse principale";
                    box.UserFields.Fields.Item(Db.F_Box_CashAcct).Value = legacyAcct;
                    box.UserFields.Fields.Item(Db.F_Box_Active).Value = "Y";
                    DiCompany.ThrowIfError(box.Add(), "Création de la caisse CAISSE1 (migration)");
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(box);
                }
            }

            string defaultBox = Scalar("SELECT TOP 1 \"Code\" FROM \"@" + Db.CashBoxTable + "\" ORDER BY \"Code\"");
            if (defaultBox == null)
                return;

            // Sessions antérieures sans caisse
            foreach (string code in Codes("SELECT \"Code\" FROM \"@" + Db.SessionTable + "\" WHERE COALESCE(\"" + Db.F_Session_CashBox + "\", '') = ''"))
            {
                UserTable table = company.UserTables.Item(Db.SessionTable);
                try
                {
                    if (!table.GetByKey(code))
                        continue;
                    table.UserFields.Fields.Item(Db.F_Session_CashBox).Value = defaultBox;
                    table.UserFields.Fields.Item(Db.F_Session_ExpOpen).Value = table.UserFields.Fields.Item(Db.F_Session_OpenBal).Value;
                    DiCompany.ThrowIfError(table.Update(), "Rattachement de la session " + code + " à la caisse " + defaultBox);
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
                }
            }

            // Verrous des sessions ouvertes (la plus récente garde la caisse)
            foreach (string code in Codes("SELECT \"Code\" FROM \"@" + Db.SessionTable + "\" WHERE \"" + Db.F_Session_Status + "\" = 'O' ORDER BY \"Code\" DESC"))
            {
                SessionRow s = LoadSession(code);
                TryAddLock(CashBoxLockKey(s.CashBox), s.Code);
                if (!string.IsNullOrEmpty(s.UserCode))
                    TryAddLock(UserLockKey(s.UserCode), s.Code);
            }
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
            string value = Scalar("SELECT \"CurrTotal\" FROM \"OACT\" WHERE \"AcctCode\" = '" + Sql(acctCode) + "'");
            if (value == null)
                throw new InvalidOperationException("Le compte G/L « " + acctCode + " » n'existe pas.");
            return Convert.ToDouble(value, CultureInfo.InvariantCulture);
        }

        private static bool BusinessPartnerExists(string cardCode)
        {
            return Scalar("SELECT \"CardCode\" FROM \"OCRD\" WHERE \"CardCode\" = '" + Sql(cardCode) + "'") != null;
        }

        /// <summary>Première colonne de la première ligne, en texte invariant (null si aucune ligne).</summary>
        private static string Scalar(string sql)
        {
            Recordset rs = NewRecordset();
            try
            {
                rs.DoQuery(sql);
                if (rs.EoF)
                    return null;
                object v = rs.Fields.Item(0).Value;
                return v is IFormattable f ? f.ToString(null, CultureInfo.InvariantCulture) : Convert.ToString(v);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }

        private static List<string> Codes(string sql)
        {
            var result = new List<string>();
            Recordset rs = NewRecordset();
            try
            {
                rs.DoQuery(sql);
                while (!rs.EoF)
                {
                    result.Add(Convert.ToString(rs.Fields.Item(0).Value));
                    rs.MoveNext();
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
            return result;
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

        /// <summary>Montant mis en forme selon le paramétrage de la société (ex. "139,500").</summary>
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
