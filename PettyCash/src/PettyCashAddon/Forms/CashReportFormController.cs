using System;
using System.Collections.Generic;
using System.Globalization;
using SAPbouiCOM;
using PettyCashAddon.Models;
using PettyCashAddon.Services;

namespace PettyCashAddon.Forms
{
    /// <summary>
    /// Écran "Rapport de caisse" : sessions détaillées ou synthèses par caisse
    /// et par utilisateur sur une période, filtrables par caisse, utilisateur
    /// et quart ; détail des opérations et du billetage de la session
    /// sélectionnée ; clôture d'une session oubliée (ouvreur ou superutilisateur).
    /// </summary>
    internal class CashReportFormController
    {
        private const string UdFrom = "udFrom";
        private const string UdTo = "udTo";
        private const string UdBoxF = "udBoxF";
        private const string UdUserF = "udUserF";
        private const string UdShiftF = "udShF";
        private const string UdView = "udView";
        private const string DtMain = "dtMain";
        private const string DtTrans = "dtTrans";
        private const string DtCount = "dtCount";

        private const string ViewSessions = "S";
        private const string ViewByBox = "C";
        private const string ViewByUser = "U";

        private static readonly Dictionary<string, string> Captions = new Dictionary<string, string>
        {
            { "Date", "Date" }, { "Caisse", "Caisse" }, { "Quart", "Quart" }, { "Util", "Utilisateur" },
            { "Ouv", "Ouverture" }, { "EcOuv", "Écart ouv." }, { "Rec", "Recettes" }, { "Dep", "Dépenses" },
            { "Theo", "Théorique" }, { "Compte", "Compté" }, { "Ecart", "Écart clôt." }, { "Statut", "Statut" },
            { "NbSess", "Sessions" }, { "NbOuv", "En cours" },
            { "Heure", "Heure" }, { "Sens", "Sens" }, { "Type", "Type" }, { "Montant", "Montant" },
            { "Tiers", "Tiers" }, { "Descr", "Description" }, { "Ecriture", "N° écriture" },
            { "Coupure", "Coupure" }, { "QtOuv", "Qté ouv." }, { "QtClo", "Qté clôt." }, { "MtClo", "Montant clôt." }
        };

        private readonly Application _app;
        private readonly BillCountForm _billForm;
        private readonly Action _onSessionClosed;
        private Form _form;
        private string _selectedSession;

        public CashReportFormController(Application app, BillCountForm billForm, Action onSessionClosed)
        {
            _app = app;
            _billForm = billForm;
            _onSessionClosed = onSessionClosed;
            _app.ItemEvent += App_ItemEvent;
        }

        public void ShowOrActivate()
        {
            if (_form != null)
            {
                _form.Select();
                return;
            }
            Build();
        }

        private void Build()
        {
            FormCreationParams p = (FormCreationParams)_app.CreateObject(BoCreatableObjectType.cot_FormCreationParams);
            p.FormType = FormIds.ReportForm;
            p.UniqueID = FormIds.ReportForm;
            p.BorderStyle = BoFormBorderStyle.fbs_Fixed;

            _form = _app.Forms.AddEx(p);
            _form.Freeze(true);
            try
            {
                _form.Title = "Rapport de caisse";
                _form.Width = 910;
                _form.Height = 640;

                // Dates liées à des sources de type date : SAP gère le format
                // (paramétrage société) et propose le calendrier.
                UserDataSources uds = _form.DataSources.UserDataSources;
                uds.Add(UdFrom, BoDataType.dt_DATE);
                uds.Add(UdTo, BoDataType.dt_DATE);
                uds.Add(UdBoxF, BoDataType.dt_SHORT_TEXT, 20);
                uds.Add(UdUserF, BoDataType.dt_SHORT_TEXT, 25);
                uds.Add(UdShiftF, BoDataType.dt_SHORT_TEXT, 1);
                uds.Add(UdView, BoDataType.dt_SHORT_TEXT, 1);

                AddLabel(FormIds.LblFrom, "Du", 10, 15, 30, FormIds.TxtFrom);
                AddEdit(FormIds.TxtFrom, 45, 12, 90, UdFrom);
                uds.Item(UdFrom).ValueEx = DateTime.Today.AddDays(-7).ToString("yyyyMMdd", CultureInfo.InvariantCulture);

                AddLabel(FormIds.LblTo, "Au", 145, 15, 25, FormIds.TxtTo);
                AddEdit(FormIds.TxtTo, 172, 12, 90, UdTo);
                uds.Item(UdTo).ValueEx = DateTime.Today.ToString("yyyyMMdd", CultureInfo.InvariantCulture);

                AddLabel(FormIds.LblBoxFilter, "Caisse", 280, 15, 45, FormIds.CmbBoxFilter);
                ComboBox cmbBox = AddCombo(FormIds.CmbBoxFilter, 330, 12, 170, UdBoxF);
                cmbBox.ValidValues.Add("*", "Toutes");
                foreach (var box in CashSessionService.GetCashBoxes(false))
                    cmbBox.ValidValues.Add(box.Code, box.Name);
                uds.Item(UdBoxF).ValueEx = "*";

                AddLabel(FormIds.LblUserFilter, "Utilisateur", 515, 15, 65, FormIds.CmbUserFilter);
                ComboBox cmbUser = AddCombo(FormIds.CmbUserFilter, 585, 12, 170, UdUserF);
                cmbUser.ValidValues.Add("*", "Tous");
                foreach (var u in CashSessionService.GetSessionUsers())
                    cmbUser.ValidValues.Add(u.Key, u.Value);
                uds.Item(UdUserF).ValueEx = "*";

                AddLabel(FormIds.LblShiftFilter, "Quart", 10, 43, 30, FormIds.CmbShiftFilter);
                ComboBox cmbShift = AddCombo(FormIds.CmbShiftFilter, 45, 40, 110, UdShiftF);
                cmbShift.ValidValues.Add("*", "Tous");
                cmbShift.ValidValues.Add("M", "Matin");
                cmbShift.ValidValues.Add("A", "Après-midi");
                cmbShift.ValidValues.Add("S", "Soir");
                uds.Item(UdShiftF).ValueEx = "*";

                AddLabel(FormIds.LblView, "Vue", 172, 43, 30, FormIds.CmbView);
                ComboBox cmbView = AddCombo(FormIds.CmbView, 205, 40, 220, UdView);
                cmbView.ValidValues.Add(ViewSessions, "Sessions détaillées");
                cmbView.ValidValues.Add(ViewByBox, "Synthèse par caisse");
                cmbView.ValidValues.Add(ViewByUser, "Synthèse par utilisateur");
                uds.Item(UdView).ValueEx = ViewSessions;

                AddButton(FormIds.BtnSearch, "Rechercher", 440, 38, 100);

                AddGrid(FormIds.GrdMain, DtMain, 10, 72, 875, 230);
                AddLabel(FormIds.LblTotals, "", 10, 308, 875, "");

                AddLabel(FormIds.LblDetail, "Opérations de la session sélectionnée", 10, 333, 400, "");
                AddGrid(FormIds.GrdTrans, DtTrans, 10, 352, 575, 205);
                AddLabel(FormIds.LblCount, "Billetage (ouverture / clôture)", 595, 333, 290, "");
                AddGrid(FormIds.GrdCount, DtCount, 595, 352, 290, 205);

                AddButton(FormIds.BtnReportClose, "Clôturer la session sélectionnée", 10, 568, 220);
            }
            catch
            {
                Form broken = _form;
                _form = null;
                try { broken.Close(); } catch { }
                throw;
            }
            finally
            {
                if (_form != null)
                    _form.Freeze(false);
            }

            _form.Visible = true;

            try
            {
                RunSearch();
            }
            catch (Exception ex)
            {
                _app.MessageBox(ex.Message);
            }
        }

        private void App_ItemEvent(string formUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            if (formUID != FormIds.ReportForm)
                return;

            if (pVal.EventType == BoEventTypes.et_FORM_UNLOAD && !pVal.BeforeAction)
            {
                _form = null;
                _selectedSession = null;
                return;
            }

            if (pVal.BeforeAction)
                return;

            try
            {
                if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.ActionSuccess)
                {
                    if (pVal.ItemUID == FormIds.BtnSearch)
                        RunSearch();
                    else if (pVal.ItemUID == FormIds.BtnReportClose)
                        HandleClose();
                    return;
                }

                if (pVal.EventType == BoEventTypes.et_COMBO_SELECT && pVal.ItemUID == FormIds.CmbView)
                {
                    RunSearch();
                    return;
                }

                if (pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == FormIds.GrdMain && pVal.Row >= 0 && CurrentView() == ViewSessions)
                    SelectSessionRow(pVal.Row);
            }
            catch (Exception ex)
            {
                _app.MessageBox(ex.Message);
            }
        }

        private string CurrentView()
        {
            string v = _form.DataSources.UserDataSources.Item(UdView).ValueEx;
            return string.IsNullOrEmpty(v) ? ViewSessions : v;
        }

        private void RunSearch()
        {
            UserDataSources uds = _form.DataSources.UserDataSources;

            if (!TryParseDate(uds.Item(UdFrom).ValueEx, out DateTime from) ||
                !TryParseDate(uds.Item(UdTo).ValueEx, out DateTime to))
            {
                _app.MessageBox("Renseignez les dates de début et de fin.");
                return;
            }
            if (from > to)
            {
                _app.MessageBox("La date de début doit être antérieure à la date de fin.");
                return;
            }

            string box = Filter(uds.Item(UdBoxF).ValueEx);
            string user = Filter(uds.Item(UdUserF).ValueEx);
            string shiftCode = Filter(uds.Item(UdShiftF).ValueEx);
            Shift? shift = shiftCode == null ? (Shift?)null : EnumCodes.ShiftFromCode(shiftCode);
            string view = CurrentView();

            string sql = view == ViewSessions
                ? CashSessionService.SessionsReportSql(from, to, box, user, shift)
                : CashSessionService.SummaryReportSql(view == ViewByBox, from, to, box, user, shift);

            DataTable dt;
            _form.Freeze(true);
            try
            {
                dt = LoadGrid(FormIds.GrdMain, DtMain, sql);
                if (view == ViewSessions)
                    ((Grid)_form.Items.Item(FormIds.GrdMain).Specific).Columns.Item("Code").Visible = false;

                ShowTotals(dt, view);
                ClearDetail();
            }
            finally
            {
                _form.Freeze(false);
            }

            int count = IsEmpty(dt) ? 0 : dt.Rows.Count;
            _app.StatusBar.SetText(count + " ligne(s).", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
        }

        private void ShowTotals(DataTable dt, string view)
        {
            double rec = 0, dep = 0, ecOuv = 0, ecart = 0;
            int sessions = 0;
            if (!IsEmpty(dt))
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    rec += Convert.ToDouble(dt.GetValue("Rec", i));
                    dep += Convert.ToDouble(dt.GetValue("Dep", i));
                    ecOuv += Convert.ToDouble(dt.GetValue("EcOuv", i));
                    ecart += Convert.ToDouble(dt.GetValue("Ecart", i));
                    sessions += view == ViewSessions ? 1 : Convert.ToInt32(dt.GetValue("NbSess", i));
                }
            }

            ((StaticText)_form.Items.Item(FormIds.LblTotals).Specific).Caption =
                "Total : " + sessions + " session(s) — Recettes " + CashSessionService.FormatAmount(rec) +
                " — Dépenses " + CashSessionService.FormatAmount(dep) +
                " — Écarts d'ouverture " + CashSessionService.FormatAmount(ecOuv) +
                " — Écarts de clôture " + CashSessionService.FormatAmount(ecart);
        }

        private void SelectSessionRow(int gridRow)
        {
            Grid grid = (Grid)_form.Items.Item(FormIds.GrdMain).Specific;
            DataTable dt = _form.DataSources.DataTables.Item(DtMain);
            if (IsEmpty(dt))
                return;

            int dtRow = grid.GetDataTableRowIndex(gridRow);
            if (dtRow < 0 || dtRow >= dt.Rows.Count)
                return;

            grid.Rows.SelectedRows.Clear();
            grid.Rows.SelectedRows.Add(gridRow);

            _selectedSession = Convert.ToString(dt.GetValue("Code", dtRow));
            _form.Freeze(true);
            try
            {
                LoadGrid(FormIds.GrdTrans, DtTrans, CashSessionService.TransactionsReportSql(_selectedSession));
                LoadGrid(FormIds.GrdCount, DtCount, CashSessionService.CountReportSql(_selectedSession));
            }
            finally
            {
                _form.Freeze(false);
            }

            SessionRow s = CashSessionService.LoadSession(_selectedSession);
            _form.Items.Item(FormIds.BtnReportClose).Enabled = CashSessionService.CanClose(s);
        }

        private void ClearDetail()
        {
            _selectedSession = null;
            // Requêtes sans résultat : vide les grilles en gardant leurs colonnes
            LoadGrid(FormIds.GrdTrans, DtTrans, CashSessionService.TransactionsReportSql(""));
            LoadGrid(FormIds.GrdCount, DtCount, CashSessionService.CountReportSql(""));
            _form.Items.Item(FormIds.BtnReportClose).Enabled = false;
        }

        private void HandleClose()
        {
            if (_selectedSession == null)
            {
                _app.MessageBox("Sélectionnez d'abord une session ouverte dans la liste.");
                return;
            }
            SessionRow session = CashSessionService.LoadSession(_selectedSession);
            if (session == null || session.Status != SessionStatus.Open)
            {
                _app.MessageBox("Cette session est déjà clôturée.");
                return;
            }

            CashSessionFormController.ShowCloseCount(_app, _billForm, session, () =>
            {
                RunSearch();
                _onSessionClosed?.Invoke();
            });
        }

        // ---------- Helpers ----------

        private DataTable LoadGrid(string gridId, string dtId, string sql)
        {
            DataTable dt = _form.DataSources.DataTables.Item(dtId);
            dt.ExecuteQuery(sql);
            Grid grid = (Grid)_form.Items.Item(gridId).Specific;
            grid.DataTable = dt;
            for (int i = 0; i < grid.Columns.Count; i++)
            {
                GridColumn col = grid.Columns.Item(i);
                col.Editable = false;
                if (Captions.TryGetValue(col.UniqueID, out string caption))
                    col.TitleObject.Caption = caption;
            }
            grid.AutoResizeColumns();
            return dt;
        }

        private static bool IsEmpty(DataTable dt)
        {
            return dt.IsEmpty || dt.Rows.Count == 0;
        }

        private static string Filter(string value)
        {
            return string.IsNullOrEmpty(value) || value == "*" ? null : value;
        }

        private static bool TryParseDate(string valueEx, out DateTime date)
        {
            // ValueEx d'une source de type date : "yyyyMMdd"
            return DateTime.TryParseExact(valueEx, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out date);
        }

        private void AddGrid(string id, string dtId, int left, int top, int width, int height)
        {
            _form.DataSources.DataTables.Add(dtId);
            Item item = _form.Items.Add(id, BoFormItemTypes.it_GRID);
            item.Left = left; item.Top = top; item.Width = width; item.Height = height;
            ((Grid)item.Specific).SelectionMode = BoMatrixSelect.ms_Single;
        }

        private void AddLabel(string id, string caption, int left, int top, int width, string linkTo)
        {
            Item item = _form.Items.Add(id, BoFormItemTypes.it_STATIC);
            item.Left = left; item.Top = top; item.Width = width; item.Height = 15;
            if (!string.IsNullOrEmpty(linkTo))
                item.LinkTo = linkTo;
            ((StaticText)item.Specific).Caption = caption;
        }

        private void AddEdit(string id, int left, int top, int width, string udsId)
        {
            Item item = _form.Items.Add(id, BoFormItemTypes.it_EDIT);
            item.Left = left; item.Top = top; item.Width = width; item.Height = 19;
            ((EditText)item.Specific).DataBind.SetBound(true, "", udsId);
        }

        private ComboBox AddCombo(string id, int left, int top, int width, string udsId)
        {
            Item item = _form.Items.Add(id, BoFormItemTypes.it_COMBO_BOX);
            item.Left = left; item.Top = top; item.Width = width; item.Height = 19;
            item.DisplayDesc = true;
            ComboBox cmb = (ComboBox)item.Specific;
            cmb.DataBind.SetBound(true, "", udsId);
            return cmb;
        }

        private void AddButton(string id, string caption, int left, int top, int width)
        {
            Item item = _form.Items.Add(id, BoFormItemTypes.it_BUTTON);
            item.Left = left; item.Top = top; item.Width = width; item.Height = 22;
            ((Button)item.Specific).Caption = caption;
        }
    }
}
