namespace PettyCashAddon
{
    /// <summary>
    /// Noms de tables et de champs UDT. Centralisés ici pour éviter les
    /// fautes de frappe dans le reste du code — c'est la seule source de
    /// vérité sur le schéma de données.
    /// Limites SAP : nom de table ≤ 19 car., nom de champ ≤ 18 car.,
    /// descriptions ≤ 30 car. ; index unique sur Code ET sur Name.
    /// </summary>
    internal static class Db
    {
        // Table de paramétrage (une seule ligne, Code = "1")
        public const string SettingsTable = "PC_SETTINGS";
        public const string SettingsCode = "1";
        public const string F_CashAcct = "U_CashAcct";   // historique : sert à créer la 1re caisse (migration)
        public const string F_DiffAcct = "U_DiffAcct";   // compte d'écarts par défaut de toutes les caisses

        // Caisses (Code = identifiant court, Name = libellé affiché, unique)
        public const string CashBoxTable = "PC_CASHBOX";
        public const string F_Box_CashAcct = "U_CashAcct";
        public const string F_Box_DiffAcct = "U_DiffAcct";   // facultatif : remplace le compte d'écarts par défaut
        public const string F_Box_Active = "U_Active";       // 'Y' / 'N'

        // Table des types d'opération (Recette/Dépense -> compte G/L)
        public const string TTypeTable = "PC_TTYPE";
        public const string F_TType_Dir = "U_Dir";       // 'R' ou 'D'
        public const string F_TType_GLAcct = "U_GLAcct";

        // Table des sessions de caisse (une par caisse et par quart)
        public const string SessionTable = "PC_SESSION";
        public const string F_Session_Date = "U_CashDate";
        public const string F_Session_Shift = "U_Shift";     // 'M' / 'A' / 'S'
        public const string F_Session_Cashier = "U_Cashier"; // nom de l'utilisateur SAP (affichage)
        public const string F_Session_User = "U_UserCode";   // code utilisateur SAP qui a ouvert la session
        public const string F_Session_CashBox = "U_CashBox";
        public const string F_Session_ExpOpen = "U_ExpOpen"; // solde d'ouverture attendu (repris)
        public const string F_Session_OpenDiff = "U_OpenDiff"; // écart constaté au billetage d'ouverture
        public const string F_Session_OpenBal = "U_OpenBal"; // solde d'ouverture = billetage d'ouverture
        public const string F_Session_TheoBal = "U_TheoBal";
        public const string F_Session_CountBal = "U_CountBal";
        public const string F_Session_Diff = "U_Diff";
        public const string F_Session_Status = "U_Status";   // 'O' Ouverte / 'C' Clôturée
        public const string F_Session_ClosedAt = "U_ClosedAt";
        public const string F_Session_ClosedBy = "U_ClosedBy";

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

        // Coupures (billets et pièces) proposées au billetage
        public const string DenomTable = "PC_DENOM";
        public const string F_Denom_Value = "U_Value";
        public const string F_Denom_Kind = "U_Kind";         // 'B' billet / 'P' pièce
        public const string F_Denom_Active = "U_Active";     // 'Y' / 'N'

        // Lignes de billetage d'une session
        public const string CountTable = "PC_COUNT";
        public const string F_Count_Session = "U_SessCode";
        public const string F_Count_Phase = "U_Phase";       // 'O' ouverture / 'C' clôture
        public const string F_Count_Denom = "U_Denom";
        public const string F_Count_Value = "U_Value";
        public const string F_Count_Qty = "U_Qty";
        public const string F_Count_Amount = "U_Amount";

        // Verrous : une ligne par caisse occupée ("CB:<caisse>") et par utilisateur
        // ayant une session ouverte ("US:<utilisateur>"). L'index unique de SAP sur
        // Code garantit qu'une caisse ne peut pas être ouverte deux fois, même par
        // deux postes au même instant.
        public const string LockTable = "PC_LOCK";
        public const string F_Lock_Session = "U_SessCode";
    }
}
