using System;
using SAPbobsCOM;
using PettyCashAddon.Core;

namespace PettyCashAddon.Setup
{
    /// <summary>
    /// Crée au premier démarrage (et uniquement s'ils n'existent pas déjà)
    /// les tables et champs utilisateur nécessaires à l'add-on. Idempotent :
    /// peut être appelé à chaque démarrage sans risque.
    /// </summary>
    internal static class MetadataSetup
    {
        public static void EnsureSchema()
        {
            Company company = DiCompany.Instance;

            EnsureTable(company, Db.SettingsTable, "Petty Cash - Paramètres");
            EnsureField(company, Db.SettingsTable, Db.F_CashAcct, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 20, "Compte G/L Caisse");
            EnsureField(company, Db.SettingsTable, Db.F_DiffAcct, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 20, "Compte G/L Écarts de caisse");
            EnsureSettingsRow(company);

            EnsureTable(company, Db.TTypeTable, "Petty Cash - Types d'opération");
            EnsureField(company, Db.TTypeTable, Db.F_TType_Dir, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "Sens (R/D)");
            EnsureField(company, Db.TTypeTable, Db.F_TType_GLAcct, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 20, "Compte G/L");

            EnsureTable(company, Db.CashBoxTable, "Petty Cash - Caisses");
            EnsureField(company, Db.CashBoxTable, Db.F_Box_CashAcct, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 20, "Compte G/L Caisse");
            EnsureField(company, Db.CashBoxTable, Db.F_Box_DiffAcct, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 20, "Compte G/L Écarts (option)");
            EnsureField(company, Db.CashBoxTable, Db.F_Box_Active, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "Active (Y/N)");

            EnsureTable(company, Db.SessionTable, "Petty Cash - Sessions");
            EnsureField(company, Db.SessionTable, Db.F_Session_Date, BoFieldTypes.db_Date, BoFldSubTypes.st_None, 0, "Date");
            EnsureField(company, Db.SessionTable, Db.F_Session_Shift, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "Quart (M/A/S)");
            EnsureField(company, Db.SessionTable, Db.F_Session_Cashier, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 50, "Caissier");
            EnsureField(company, Db.SessionTable, Db.F_Session_User, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 25, "Utilisateur SAP");
            EnsureField(company, Db.SessionTable, Db.F_Session_CashBox, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 20, "Caisse");
            EnsureField(company, Db.SessionTable, Db.F_Session_ExpOpen, BoFieldTypes.db_Float, BoFldSubTypes.st_Sum, 0, "Ouverture attendue");
            EnsureField(company, Db.SessionTable, Db.F_Session_OpenDiff, BoFieldTypes.db_Float, BoFldSubTypes.st_Sum, 0, "Écart d'ouverture");
            EnsureField(company, Db.SessionTable, Db.F_Session_OpenBal, BoFieldTypes.db_Float, BoFldSubTypes.st_Sum, 0, "Solde ouverture");
            EnsureField(company, Db.SessionTable, Db.F_Session_TheoBal, BoFieldTypes.db_Float, BoFldSubTypes.st_Sum, 0, "Solde théorique");
            EnsureField(company, Db.SessionTable, Db.F_Session_CountBal, BoFieldTypes.db_Float, BoFldSubTypes.st_Sum, 0, "Solde compté");
            EnsureField(company, Db.SessionTable, Db.F_Session_Diff, BoFieldTypes.db_Float, BoFldSubTypes.st_Sum, 0, "Écart");
            EnsureField(company, Db.SessionTable, Db.F_Session_Status, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "Statut (O/C)");
            EnsureField(company, Db.SessionTable, Db.F_Session_ClosedAt, BoFieldTypes.db_Date, BoFldSubTypes.st_None, 0, "Date de clôture");
            EnsureField(company, Db.SessionTable, Db.F_Session_ClosedBy, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 25, "Clôturée par");

            EnsureTable(company, Db.DenomTable, "Petty Cash - Coupures");
            EnsureField(company, Db.DenomTable, Db.F_Denom_Value, BoFieldTypes.db_Float, BoFldSubTypes.st_Sum, 0, "Valeur");
            EnsureField(company, Db.DenomTable, Db.F_Denom_Kind, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "Billet/Pièce (B/P)");
            EnsureField(company, Db.DenomTable, Db.F_Denom_Active, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "Active (Y/N)");
            SeedDenominations(company);

            EnsureTable(company, Db.CountTable, "Petty Cash - Billetage");
            EnsureField(company, Db.CountTable, Db.F_Count_Session, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 30, "Session");
            EnsureField(company, Db.CountTable, Db.F_Count_Phase, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "Ouverture/Clôture (O/C)");
            EnsureField(company, Db.CountTable, Db.F_Count_Denom, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 20, "Coupure");
            EnsureField(company, Db.CountTable, Db.F_Count_Value, BoFieldTypes.db_Float, BoFldSubTypes.st_Sum, 0, "Valeur");
            EnsureField(company, Db.CountTable, Db.F_Count_Qty, BoFieldTypes.db_Numeric, BoFldSubTypes.st_None, 11, "Quantité");
            EnsureField(company, Db.CountTable, Db.F_Count_Amount, BoFieldTypes.db_Float, BoFldSubTypes.st_Sum, 0, "Montant");

            EnsureTable(company, Db.LockTable, "Petty Cash - Verrous");
            EnsureField(company, Db.LockTable, Db.F_Lock_Session, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 30, "Session");

            EnsureTable(company, Db.TransTable, "Petty Cash - Transactions");
            EnsureField(company, Db.TransTable, Db.F_Trans_Session, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 30, "Session");
            EnsureField(company, Db.TransTable, Db.F_Trans_Time, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 8, "Heure");
            EnsureField(company, Db.TransTable, Db.F_Trans_Dir, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "Sens (R/D)");
            EnsureField(company, Db.TransTable, Db.F_Trans_TType, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 20, "Type d'opération");
            EnsureField(company, Db.TransTable, Db.F_Trans_Amount, BoFieldTypes.db_Float, BoFldSubTypes.st_Sum, 0, "Montant");
            EnsureField(company, Db.TransTable, Db.F_Trans_CardCode, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 15, "Tiers");
            EnsureField(company, Db.TransTable, Db.F_Trans_Descript, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, "Description");
            EnsureField(company, Db.TransTable, Db.F_Trans_JE, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 15, "N° écriture");

            // Passage mono-caisse -> multi-caisses (sans effet si déjà fait)
            Services.CashSessionService.MigrateToMultiCashBox();
        }

        /// <summary>Coupures du franc CFA, créées seulement si la table est vide.</summary>
        private static void SeedDenominations(Company company)
        {
            Recordset rs = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                rs.DoQuery("SELECT COUNT(*) FROM \"@" + Db.DenomTable + "\"");
                if (Convert.ToInt32(rs.Fields.Item(0).Value) > 0)
                    return;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }

            var denoms = new (string Kind, int Value)[]
            {
                ("B", 10000), ("B", 5000), ("B", 2000), ("B", 1000), ("B", 500),
                ("P", 500), ("P", 250), ("P", 200), ("P", 100), ("P", 50), ("P", 25), ("P", 10), ("P", 5)
            };
            foreach (var d in denoms)
            {
                UserTable table = company.UserTables.Item(Db.DenomTable);
                try
                {
                    table.Code = d.Kind + d.Value;
                    table.Name = (d.Kind == "B" ? "Billet " : "Pièce ") + d.Value.ToString("N0", System.Globalization.CultureInfo.GetCultureInfo("fr-FR"));
                    table.UserFields.Fields.Item(Db.F_Denom_Value).Value = (double)d.Value;
                    table.UserFields.Fields.Item(Db.F_Denom_Kind).Value = d.Kind;
                    table.UserFields.Fields.Item(Db.F_Denom_Active).Value = "Y";
                    DiCompany.ThrowIfError(table.Add(), "Création de la coupure " + d.Kind + d.Value);
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
                }
            }
        }

        private static void EnsureTable(Company company, string tableName, string description)
        {
            UserTablesMD tablesMd = (UserTablesMD)company.GetBusinessObject(BoObjectTypes.oUserTables);
            try
            {
                if (tablesMd.GetByKey(tableName))
                    return;

                tablesMd.TableName = tableName;
                tablesMd.TableDescription = description;
                tablesMd.TableType = BoUTBTableType.bott_NoObject;

                int rc = tablesMd.Add();
                DiCompany.ThrowIfError(rc, "Création de la table @" + tableName);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(tablesMd);
            }
        }

        private static void EnsureField(Company company, string tableName, string fieldName, BoFieldTypes type, BoFldSubTypes subType, int size, string description)
        {
            string userFieldName = fieldName.StartsWith("U_") ? fieldName.Substring(2) : fieldName;

            if (FieldExists(company, tableName, userFieldName))
                return;

            UserFieldsMD fieldsMd = (UserFieldsMD)company.GetBusinessObject(BoObjectTypes.oUserFields);
            try
            {
                fieldsMd.TableName = "@" + tableName;
                fieldsMd.Name = userFieldName;
                fieldsMd.Type = type;
                fieldsMd.SubType = subType;
                if (size > 0)
                    fieldsMd.EditSize = size;
                fieldsMd.Description = description;

                int rc = fieldsMd.Add();
                DiCompany.ThrowIfError(rc, "Création du champ " + fieldName + " sur @" + tableName);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(fieldsMd);
            }
        }

        private static bool FieldExists(Company company, string tableName, string userFieldName)
        {
            Recordset rs = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                string sql = "SELECT \"FieldID\" FROM \"CUFD\" WHERE \"TableID\" = '@" + tableName + "' AND \"AliasID\" = '" + userFieldName + "'";
                rs.DoQuery(sql);
                return !rs.EoF;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }

        private static void EnsureSettingsRow(Company company)
        {
            UserTable table = company.UserTables.Item(Db.SettingsTable);
            try
            {
                if (table.GetByKey(Db.SettingsCode))
                    return;

                table.Code = Db.SettingsCode;
                table.Name = "Paramètres caisse";
                int rc = table.Add();
                DiCompany.ThrowIfError(rc, "Initialisation de la ligne de paramétrage");
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
            }
        }
    }
}
