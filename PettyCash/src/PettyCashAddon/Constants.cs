namespace PettyCashAddon
{
    /// <summary>
    /// Noms de tables et de champs UDT. Centralisés ici pour éviter les
    /// fautes de frappe dans le reste du code — c'est la seule source de
    /// vérité sur le schéma de données.
    /// </summary>
    internal static class Db
    {
        // Table de paramétrage (une seule ligne, Code = "1")
        public const string SettingsTable = "PC_SETTINGS";
        public const string SettingsCode = "1";
        public const string F_CashAcct = "U_CashAcct";
        public const string F_DiffAcct = "U_DiffAcct";

        // Table des types d'opération (Recette/Dépense -> compte G/L)
        public const string TTypeTable = "PC_TTYPE";
        public const string F_TType_Dir = "U_Dir";       // 'R' ou 'D'
        public const string F_TType_GLAcct = "U_GLAcct";

        // Table des sessions de caisse (une par quart/journée)
        public const string SessionTable = "PC_SESSION";
        public const string F_Session_Date = "U_CashDate";
        public const string F_Session_Shift = "U_Shift";     // 'M' / 'A' / 'S'
        public const string F_Session_Cashier = "U_Cashier";
        public const string F_Session_OpenBal = "U_OpenBal";
        public const string F_Session_TheoBal = "U_TheoBal";
        public const string F_Session_CountBal = "U_CountBal";
        public const string F_Session_Diff = "U_Diff";
        public const string F_Session_Status = "U_Status";   // 'O' Ouverte / 'C' Clôturée
        public const string F_Session_ClosedAt = "U_ClosedAt";

        // Table des transactions de caisse (lignes d'une session)
        public const string TransTable = "PC_TRANS";
        public const string F_Trans_Session = "U_SessCode";
        public const string F_Trans_Time = "U_Time";
        public const string F_Trans_Dir = "U_Dir";           // 'R' ou 'D'
        public const string F_Trans_TType = "U_TType";
        public const string F_Trans_Amount = "U_Amount";
        public const string F_Trans_CardCode = "U_CardCode";
        public const string F_Trans_Descript = "U_Descript";
        public const string F_Trans_JE = "U_JE";
    }
}
