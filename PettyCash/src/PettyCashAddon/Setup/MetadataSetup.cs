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
            EnsureField(company, Db.SettingsTable, Db.F_CashAcct, BoFieldTypes.db_Alpha, 20, "Compte G/L Caisse");
            EnsureField(company, Db.SettingsTable, Db.F_DiffAcct, BoFieldTypes.db_Alpha, 20, "Compte G/L Écarts de caisse");
            EnsureSettingsRow(company);

            EnsureTable(company, Db.TTypeTable, "Petty Cash - Types d'opération");
            EnsureField(company, Db.TTypeTable, Db.F_TType_Dir, BoFieldTypes.db_Alpha, 1, "Sens (R/D)");
            EnsureField(company, Db.TTypeTable, Db.F_TType_GLAcct, BoFieldTypes.db_Alpha, 20, "Compte G/L");

            EnsureTable(company, Db.SessionTable, "Petty Cash - Sessions de caisse");
            EnsureField(company, Db.SessionTable, Db.F_Session_Date, BoFieldTypes.db_Date, 0, "Date");
            EnsureField(company, Db.SessionTable, Db.F_Session_Shift, BoFieldTypes.db_Alpha, 1, "Quart (M/A/S)");
            EnsureField(company, Db.SessionTable, Db.F_Session_Cashier, BoFieldTypes.db_Alpha, 50, "Caissier");
            EnsureField(company, Db.SessionTable, Db.F_Session_OpenBal, BoFieldTypes.db_Float, 0, "Solde ouverture");
            EnsureField(company, Db.SessionTable, Db.F_Session_TheoBal, BoFieldTypes.db_Float, 0, "Solde théorique");
            EnsureField(company, Db.SessionTable, Db.F_Session_CountBal, BoFieldTypes.db_Float, 0, "Solde compté");
            EnsureField(company, Db.SessionTable, Db.F_Session_Diff, BoFieldTypes.db_Float, 0, "Écart");
            EnsureField(company, Db.SessionTable, Db.F_Session_Status, BoFieldTypes.db_Alpha, 1, "Statut (O/C)");
            EnsureField(company, Db.SessionTable, Db.F_Session_ClosedAt, BoFieldTypes.db_Date, 0, "Date de clôture");

            EnsureTable(company, Db.TransTable, "Petty Cash - Transactions de caisse");
            EnsureField(company, Db.TransTable, Db.F_Trans_Session, BoFieldTypes.db_Alpha, 30, "Session");
            EnsureField(company, Db.TransTable, Db.F_Trans_Time, BoFieldTypes.db_Alpha, 8, "Heure");
            EnsureField(company, Db.TransTable, Db.F_Trans_Dir, BoFieldTypes.db_Alpha, 1, "Sens (R/D)");
            EnsureField(company, Db.TransTable, Db.F_Trans_TType, BoFieldTypes.db_Alpha, 20, "Type d'opération");
            EnsureField(company, Db.TransTable, Db.F_Trans_Amount, BoFieldTypes.db_Float, 0, "Montant");
            EnsureField(company, Db.TransTable, Db.F_Trans_CardCode, BoFieldTypes.db_Alpha, 15, "Tiers");
            EnsureField(company, Db.TransTable, Db.F_Trans_Descript, BoFieldTypes.db_Alpha, 100, "Description");
            EnsureField(company, Db.TransTable, Db.F_Trans_JE, BoFieldTypes.db_Alpha, 15, "N° écriture");
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

        private static void EnsureField(Company company, string tableName, string fieldName, BoFieldTypes type, int size, string description)
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
            UserTable table = (UserTable)company.GetBusinessObject(BoObjectTypes.oUserTable);
            try
            {
                table.TableName = Db.SettingsTable;
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
