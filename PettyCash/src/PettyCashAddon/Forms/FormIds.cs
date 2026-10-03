namespace PettyCashAddon.Forms
{
    /// <summary>
    /// Identifiants des écrans et éléments. SAP limite les identifiants
    /// d'éléments (items, colonnes) à 10 caractères.
    /// </summary>
    internal static class FormIds
    {
        public const string SessionForm = "PC_SESSION_FORM";
        public const string ReportForm = "PC_REPORT_FORM";
        public const string TransactionEntryForm = "PC_TRANS_ENTRY_FORM";
        public const string BillForm = "PC_BILL_FORM";

        public const string MenuRoot = "PC_MENU_ROOT";
        public const string MenuSession = "PC_MENU_SESSION";
        public const string MenuReport = "PC_MENU_REPORT";

        // Items - écran Session
        public const string LblBox = "lblBox";
        public const string CmbBox = "cmbBox";
        public const string LblShift = "lblShift";
        public const string CmbShift = "cmbShift";
        public const string LblUser = "lblUser";
        public const string TxtUser = "txtUser";
        public const string LblStatus = "lblStatus";
        public const string TxtStatus = "txtStatus";
        public const string LblOpenBal = "lblOpenBal";
        public const string TxtOpenBal = "txtOpenBal";
        public const string LblTheoBal = "lblTheoBal";
        public const string TxtTheoBal = "txtTheoBal";
        public const string LblCountBal = "lblCount";
        public const string TxtCountBal = "txtCount";
        public const string LblDiff = "lblDiff";
        public const string TxtDiff = "txtDiff";
        public const string MatTrans = "matTrans";
        public const string BtnOpen = "btnOpen";
        public const string BtnAddTrans = "btnAddTr";
        public const string BtnClose = "btnClose";
        public const string BtnRefresh = "btnRefresh";

        // Items - écran popup transaction
        public const string LblDir = "lblDir";
        public const string CmbDir = "cmbDir";
        public const string LblType = "lblType";
        public const string CmbType = "cmbType";
        public const string LblAmount = "lblAmount";
        public const string TxtAmount = "txtAmount";
        public const string LblCard = "lblCard";
        public const string TxtCard = "txtCard";
        public const string LblDescr = "lblDescr";
        public const string TxtDescr = "txtDescr";
        public const string BtnOk = "btnOk";
        public const string BtnCancel = "btnCancel";

        // Items - écran billetage
        public const string LblBillInfo = "lblInfo";
        public const string LblExpected = "lblExp";
        public const string TxtExpected = "txtExp";
        public const string MatBill = "matBill";
        public const string LblTotal = "lblTotal";
        public const string TxtTotal = "txtTotal";
        public const string LblBillDiff = "lblBDiff";
        public const string TxtBillDiff = "txtBDiff";
        public const string BtnBillOk = "btnBOk";
        public const string BtnBillCancel = "btnBCan";

        // Items - écran rapport
        public const string LblFrom = "lblFrom";
        public const string TxtFrom = "txtFrom";
        public const string LblTo = "lblTo";
        public const string TxtTo = "txtTo";
        public const string LblBoxFilter = "lblBoxF";
        public const string CmbBoxFilter = "cmbBoxF";
        public const string LblUserFilter = "lblUserF";
        public const string CmbUserFilter = "cmbUserF";
        public const string LblShiftFilter = "lblShiftF";
        public const string CmbShiftFilter = "cmbShiftF";
        public const string LblView = "lblView";
        public const string CmbView = "cmbView";
        public const string BtnSearch = "btnSearch";
        public const string GrdMain = "grdMain";
        public const string LblTotals = "lblTotals";
        public const string LblDetail = "lblDetail";
        public const string GrdTrans = "grdTrans";
        public const string LblCount = "lblCnt";
        public const string GrdCount = "grdCount";
        public const string BtnReportClose = "btnRClose";
    }
}
