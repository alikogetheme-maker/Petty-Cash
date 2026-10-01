using System;
using System.Collections.Generic;
using System.Globalization;
using SAPbouiCOM;
using PettyCashAddon.Models;
using PettyCashAddon.Services;

namespace PettyCashAddon.Forms
{
    /// <summary>
    /// Écran "Rapport de caisse" : liste des sessions sur une période
    /// (filtrable par quart), avec détail des transactions de la session
    /// sélectionnée.
    /// </summary>
    internal class CashReportFormController
    {
        private const string UdFrom = "udFrom";
        private const string UdTo = "udTo";
        private const string UdShiftFilter = "udShF";

        private readonly Application _app;
        private Form _form;
        private List<SessionRow> _currentSessions = new List<SessionRow>();

        public CashReportFormController(Application app)
        {
            _app = app;
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
                _form.Width = 700;
                _form.Height = 500;

                // Dates liées à des sources de type date : SAP gère le format
                // (paramétrage société) et propose le calendrier.
                UserDataSources uds = _form.DataSources.UserDataSources;
                uds.Add(UdFrom, BoDataType.dt_DATE);
                uds.Add(UdTo, BoDataType.dt_DATE);
                uds.Add(UdShiftFilter, BoDataType.dt_SHORT_TEXT, 1);

                AddLabel(FormIds.LblFrom, "Du", 10, 15, 40, FormIds.TxtFrom);
                AddEdit(FormIds.TxtFrom, 55, 12, 100, UdFrom);
                uds.Item(UdFrom).ValueEx = DateTime.Today.AddDays(-7).ToString("yyyyMMdd", CultureInfo.InvariantCulture);

                AddLabel(FormIds.LblTo, "Au", 170, 15, 25, FormIds.TxtTo);
                AddEdit(FormIds.TxtTo, 200, 12, 100, UdTo);
                uds.Item(UdTo).ValueEx = DateTime.Today.ToString("yyyyMMdd", CultureInfo.InvariantCulture);

                AddLabel(FormIds.LblShiftFilter, "Quart", 320, 15, 40, FormIds.CmbShiftFilter);
                ComboBox cmbShift = AddCombo(FormIds.CmbShiftFilter, 365, 12, 130);
                cmbShift.DataBind.SetBound(true, "", UdShiftFilter);
                cmbShift.ValidValues.Add("*", "Tous");
                cmbShift.ValidValues.Add("M", "Matin");
                cmbShift.ValidValues.Add("A", "Après-midi");
                cmbShift.ValidValues.Add("S", "Soir");
                uds.Item(UdShiftFilter).ValueEx = "*";

                AddButton(FormIds.BtnSearch, "Rechercher", 510, 10, 100);

                BuildSessionsMatrix();
                BuildTransMatrix();
            }
            catch
            {
                // Écran à moitié construit : on le ferme pour pouvoir le recréer au prochain clic
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

        private void BuildSessionsMatrix()
        {
            DataTable dt = _form.DataSources.DataTables.Add("dtSessions");
            dt.Columns.Add("Date", BoFieldsType.ft_AlphaNumeric, 10);
            dt.Columns.Add("Quart", BoFieldsType.ft_AlphaNumeric, 15);
            dt.Columns.Add("Caissier", BoFieldsType.ft_AlphaNumeric, 50);
            dt.Columns.Add("Ouverture", BoFieldsType.ft_Sum);
            dt.Columns.Add("Theorique", BoFieldsType.ft_Sum);
            dt.Columns.Add("Compte", BoFieldsType.ft_Sum);
            dt.Columns.Add("Ecart", BoFieldsType.ft_Sum);
            dt.Columns.Add("Statut", BoFieldsType.ft_AlphaNumeric, 15);

            Item item = _form.Items.Add(FormIds.MatSessions, BoFormItemTypes.it_MATRIX);
            item.Left = 10; item.Top = 45; item.Width = 670; item.Height = 190;
            Matrix mat = (Matrix)item.Specific;
            mat.SelectionMode = BoMatrixSelect.ms_Single;

            AddCol(mat, "colDate", "Date", 75, "dtSessions", "Date");
            AddCol(mat, "colQuart", "Quart", 80, "dtSessions", "Quart");
            AddCol(mat, "colCaiss", "Caissier", 120, "dtSessions", "Caissier");
            AddCol(mat, "colOuv", "Ouverture", 80, "dtSessions", "Ouverture");
            AddCol(mat, "colTheo", "Théorique", 80, "dtSessions", "Theorique");
            AddCol(mat, "colCpt", "Compté", 80, "dtSessions", "Compte");
            AddCol(mat, "colEcart", "Écart", 70, "dtSessions", "Ecart");
            AddCol(mat, "colStatut", "Statut", 80, "dtSessions", "Statut");
        }

        private void BuildTransMatrix()
        {
            DataTable dt = _form.DataSources.DataTables.Add("dtSessTr");
            dt.Columns.Add("Heure", BoFieldsType.ft_AlphaNumeric, 8);
            dt.Columns.Add("Sens", BoFieldsType.ft_AlphaNumeric, 10);
            dt.Columns.Add("Type", BoFieldsType.ft_AlphaNumeric, 100);
            dt.Columns.Add("Montant", BoFieldsType.ft_Sum);
            dt.Columns.Add("Tiers", BoFieldsType.ft_AlphaNumeric, 15);
            dt.Columns.Add("Descr", BoFieldsType.ft_AlphaNumeric, 100);
            dt.Columns.Add("Ecriture", BoFieldsType.ft_AlphaNumeric, 15);

            AddLabel("lblDetail", "Détail de la session sélectionnée (cliquez sur une ligne ci-dessus) :", 10, 245, 400, "");

            Item item = _form.Items.Add(FormIds.MatSessionTrans, BoFormItemTypes.it_MATRIX);
            item.Left = 10; item.Top = 265; item.Width = 670; item.Height = 165;
            Matrix mat = (Matrix)item.Specific;

            AddCol(mat, "colTHeure", "Heure", 60, "dtSessTr", "Heure");
            AddCol(mat, "colTSens", "Sens", 60, "dtSessTr", "Sens");
            AddCol(mat, "colTType", "Type", 120, "dtSessTr", "Type");
            AddCol(mat, "colTMont", "Montant", 80, "dtSessTr", "Montant");
            AddCol(mat, "colTTiers", "Tiers", 70, "dtSessTr", "Tiers");
            AddCol(mat, "colTDescr", "Description", 150, "dtSessTr", "Descr");
            AddCol(mat, "colTJE", "N° écriture", 90, "dtSessTr", "Ecriture");
        }

        private void AddCol(Matrix mat, string colId, string caption, int width, string table, string boundColumn)
        {
            Column col = mat.Columns.Add(colId, BoFormItemTypes.it_EDIT);
            col.TitleObject.Caption = caption;
            col.Width = width;
            col.Editable = false;
            col.DataBind.Bind(table, boundColumn);
        }

        private void App_ItemEvent(string formUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            if (formUID != FormIds.ReportForm)
                return;

            if (pVal.EventType == BoEventTypes.et_FORM_UNLOAD && !pVal.BeforeAction)
            {
                _form = null;
                _currentSessions = new List<SessionRow>();
                return;
            }

            if (pVal.BeforeAction)
                return;

            try
            {
                if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.ActionSuccess && pVal.ItemUID == FormIds.BtnSearch)
                {
                    RunSearch();
                    return;
                }

                if (pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == FormIds.MatSessions && pVal.Row > 0)
                {
                    int rowIndex = pVal.Row - 1;
                    if (rowIndex < _currentSessions.Count)
                    {
                        ((Matrix)_form.Items.Item(FormIds.MatSessions).Specific).SelectRow(pVal.Row, true, false);
                        LoadSessionDetail(_currentSessions[rowIndex]);
                    }
                }
            }
            catch (Exception ex)
            {
                _app.MessageBox(ex.Message);
            }
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

            Shift? shiftFilter = null;
            string shiftCode = uds.Item(UdShiftFilter).ValueEx;
            if (!string.IsNullOrEmpty(shiftCode) && shiftCode != "*")
                shiftFilter = EnumCodes.ShiftFromCode(shiftCode);

            _currentSessions = CashSessionService.GetSessions(from, to, shiftFilter);

            DataTable dt = _form.DataSources.DataTables.Item("dtSessions");
            dt.Rows.Clear();
            if (_currentSessions.Count > 0)
                dt.Rows.Add(_currentSessions.Count);

            for (int i = 0; i < _currentSessions.Count; i++)
            {
                var s = _currentSessions[i];
                bool closed = s.Status == SessionStatus.Closed;
                dt.SetValue("Date", i, s.CashDate.ToString("dd/MM/yyyy"));
                dt.SetValue("Quart", i, ShiftLabel(s.Shift));
                dt.SetValue("Caissier", i, s.Cashier);
                dt.SetValue("Ouverture", i, s.OpenBal);
                dt.SetValue("Theorique", i, s.TheoBal);
                dt.SetValue("Compte", i, closed ? s.CountBal : 0);
                dt.SetValue("Ecart", i, closed ? s.Diff : 0);
                dt.SetValue("Statut", i, closed ? "Clôturée" : "Ouverte");
            }

            ((Matrix)_form.Items.Item(FormIds.MatSessions).Specific).LoadFromDataSource();
            ClearTransMatrix();

            _app.StatusBar.SetText(_currentSessions.Count + " session(s) trouvée(s).", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
        }

        private static bool TryParseDate(string valueEx, out DateTime date)
        {
            // ValueEx d'une source de type date : "yyyyMMdd"
            return DateTime.TryParseExact(valueEx, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out date);
        }

        private void LoadSessionDetail(SessionRow session)
        {
            var lines = CashSessionService.GetTransactions(session.Code);

            DataTable dt = _form.DataSources.DataTables.Item("dtSessTr");
            dt.Rows.Clear();
            if (lines.Count > 0)
                dt.Rows.Add(lines.Count);

            for (int i = 0; i < lines.Count; i++)
            {
                var t = lines[i];
                dt.SetValue("Heure", i, t.Time);
                dt.SetValue("Sens", i, t.Direction == Direction.Recette ? "Recette" : "Dépense");
                dt.SetValue("Type", i, t.TransTypeName);
                dt.SetValue("Montant", i, t.Amount);
                dt.SetValue("Tiers", i, t.CardCode);
                dt.SetValue("Descr", i, t.Description);
                dt.SetValue("Ecriture", i, t.JeDocEntry);
            }

            ((Matrix)_form.Items.Item(FormIds.MatSessionTrans).Specific).LoadFromDataSource();
        }

        private void ClearTransMatrix()
        {
            _form.DataSources.DataTables.Item("dtSessTr").Rows.Clear();
            ((Matrix)_form.Items.Item(FormIds.MatSessionTrans).Specific).LoadFromDataSource();
        }

        private static string ShiftLabel(Shift shift)
        {
            switch (shift)
            {
                case Shift.ApresMidi: return "Après-midi";
                case Shift.Soir: return "Soir";
                default: return "Matin";
            }
        }

        // ---------- Helpers ----------

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

        private ComboBox AddCombo(string id, int left, int top, int width)
        {
            Item item = _form.Items.Add(id, BoFormItemTypes.it_COMBO_BOX);
            item.Left = left; item.Top = top; item.Width = width; item.Height = 19;
            item.DisplayDesc = true;
            return (ComboBox)item.Specific;
        }

        private void AddButton(string id, string caption, int left, int top, int width)
        {
            Item item = _form.Items.Add(id, BoFormItemTypes.it_BUTTON);
            item.Left = left; item.Top = top; item.Width = width; item.Height = 22;
            ((Button)item.Specific).Caption = caption;
        }
    }
}
