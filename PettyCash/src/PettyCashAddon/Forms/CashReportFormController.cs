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
        private readonly Application _app;
        private Form _form;
        private List<SessionRow> _currentSessions = new List<SessionRow>();

        public CashReportFormController(Application app)
        {
            _app = app;
        }

        public void ShowOrActivate()
        {
            foreach (Form f in _app.Forms)
            {
                if (f.TypeEx == FormIds.ReportForm)
                {
                    f.Select();
                    return;
                }
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
            _form.Title = "Rapport de caisse";
            _form.Width = 700;
            _form.Height = 480;

            AddLabel("lblFrom", "Du", 10, 15, 40);
            AddEdit(FormIds.TxtFrom, 55, 12, 100).Value = DateTime.Today.AddDays(-7).ToString("dd/MM/yyyy");

            AddLabel("lblTo", "Au", 170, 15, 25);
            AddEdit(FormIds.TxtTo, 200, 12, 100).Value = DateTime.Today.ToString("dd/MM/yyyy");

            AddLabel("lblShiftFilter", "Quart", 320, 15, 40);
            ComboBox cmbShift = AddCombo(FormIds.CmbShiftFilter, 365, 12, 130);
            cmbShift.ValidValues.Add("*", "Tous");
            cmbShift.ValidValues.Add("M", "Matin");
            cmbShift.ValidValues.Add("A", "Après-midi");
            cmbShift.ValidValues.Add("S", "Soir");
            cmbShift.Select("*", BoSearchKey.psk_ByValue);

            AddButton(FormIds.BtnSearch, "Rechercher", 510, 10, 100);

            BuildSessionsMatrix();
            BuildTransMatrix();

            _form.Visible = true;
            _app.ItemEvent += App_ItemEvent;

            RunSearch();
        }

        private void BuildSessionsMatrix()
        {
            DataTable dt = _form.DataSources.DataTables.Add("dtSessions");
            dt.Columns.Add("Date", BoFieldTypes.db_Alpha);
            dt.Columns.Add("Quart", BoFieldTypes.db_Alpha);
            dt.Columns.Add("Caissier", BoFieldTypes.db_Alpha);
            dt.Columns.Add("Ouverture", BoFieldTypes.db_Float);
            dt.Columns.Add("Theorique", BoFieldTypes.db_Float);
            dt.Columns.Add("Compte", BoFieldTypes.db_Float);
            dt.Columns.Add("Ecart", BoFieldTypes.db_Float);
            dt.Columns.Add("Statut", BoFieldTypes.db_Alpha);

            Item item = _form.Items.Add(FormIds.MatSessions, BoFormItemTypes.it_MATRIX);
            item.Left = 10; item.Top = 45; item.Width = 670; item.Height = 190;
            Matrix mat = (Matrix)item.Specific;

            AddCol(mat, "colDate", "Date", 75, "dtSessions", "Date");
            AddCol(mat, "colQuart", "Quart", 80, "dtSessions", "Quart");
            AddCol(mat, "colCaissier", "Caissier", 120, "dtSessions", "Caissier");
            AddCol(mat, "colOuv", "Ouverture", 80, "dtSessions", "Ouverture");
            AddCol(mat, "colTheo", "Théorique", 80, "dtSessions", "Theorique");
            AddCol(mat, "colCpt", "Compté", 80, "dtSessions", "Compte");
            AddCol(mat, "colEcart", "Écart", 70, "dtSessions", "Ecart");
            AddCol(mat, "colStatut", "Statut", 80, "dtSessions", "Statut");
        }

        private void BuildTransMatrix()
        {
            DataTable dt = _form.DataSources.DataTables.Add("dtSessTrans");
            dt.Columns.Add("Heure", BoFieldTypes.db_Alpha);
            dt.Columns.Add("Sens", BoFieldTypes.db_Alpha);
            dt.Columns.Add("Type", BoFieldTypes.db_Alpha);
            dt.Columns.Add("Montant", BoFieldTypes.db_Float);
            dt.Columns.Add("Tiers", BoFieldTypes.db_Alpha);
            dt.Columns.Add("Description", BoFieldTypes.db_Alpha);
            dt.Columns.Add("Ecriture", BoFieldTypes.db_Alpha);

            AddLabel("lblDetail", "Détail de la session sélectionnée :", 10, 245, 250);

            Item item = _form.Items.Add(FormIds.MatSessionTrans, BoFormItemTypes.it_MATRIX);
            item.Left = 10; item.Top = 265; item.Width = 670; item.Height = 165;
            Matrix mat = (Matrix)item.Specific;

            AddCol(mat, "colTHeure", "Heure", 60, "dtSessTrans", "Heure");
            AddCol(mat, "colTSens", "Sens", 60, "dtSessTrans", "Sens");
            AddCol(mat, "colTType", "Type", 120, "dtSessTrans", "Type");
            AddCol(mat, "colTMontant", "Montant", 80, "dtSessTrans", "Montant");
            AddCol(mat, "colTTiers", "Tiers", 70, "dtSessTrans", "Tiers");
            AddCol(mat, "colTDescr", "Description", 150, "dtSessTrans", "Description");
            AddCol(mat, "colTJE", "N° écriture", 90, "dtSessTrans", "Ecriture");
        }

        private void AddCol(Matrix mat, string colId, string caption, int width, string table, string boundColumn)
        {
            Column col = mat.Columns.Add(colId, BoFormItemTypes.it_EDIT);
            col.Title = caption;
            col.Width = width;
            col.Editable = false;
            col.DataBind.SetBound(true, table, boundColumn);
        }

        private void App_ItemEvent(string formUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            if (formUID != FormIds.ReportForm)
                return;

            if (pVal.EventType == BoEventTypes.et_FORM_CLOSE)
            {
                _app.ItemEvent -= App_ItemEvent;
                return;
            }

            if (pVal.BeforeAction)
                return;

            if (pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == FormIds.BtnSearch)
            {
                try { RunSearch(); }
                catch (Exception ex) { _app.MessageBox(ex.Message); }
                return;
            }

            if (pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == FormIds.MatSessions)
            {
                int rowIndex = pVal.Row - 1;
                if (rowIndex >= 0 && rowIndex < _currentSessions.Count)
                    LoadSessionDetail(_currentSessions[rowIndex]);
            }
        }

        private void RunSearch()
        {
            EditText txtFrom = (EditText)_form.Items.Item(FormIds.TxtFrom).Specific;
            EditText txtTo = (EditText)_form.Items.Item(FormIds.TxtTo).Specific;
            ComboBox cmbShift = (ComboBox)_form.Items.Item(FormIds.CmbShiftFilter).Specific;

            if (!DateTime.TryParse(txtFrom.Value, CultureInfo.CurrentCulture, DateTimeStyles.None, out DateTime from) ||
                !DateTime.TryParse(txtTo.Value, CultureInfo.CurrentCulture, DateTimeStyles.None, out DateTime to))
            {
                _app.MessageBox("Dates invalides (format attendu jj/mm/aaaa).");
                return;
            }

            Shift? shiftFilter = null;
            if (cmbShift.Selected != null && cmbShift.Selected.Value != "*")
                shiftFilter = EnumCodes.ShiftFromCode(cmbShift.Selected.Value);

            _currentSessions = CashSessionService.GetSessions(from, to, shiftFilter);

            DataTable dt = _form.DataSources.DataTables.Item("dtSessions");
            for (int i = dt.Rows.Count - 1; i >= 0; i--)
                dt.Rows.Remove(i);

            dt.Rows.Add(_currentSessions.Count);
            for (int i = 0; i < _currentSessions.Count; i++)
            {
                var s = _currentSessions[i];
                dt.SetValue("Date", i, s.CashDate.ToString("dd/MM/yyyy"));
                dt.SetValue("Quart", i, ShiftLabel(s.Shift));
                dt.SetValue("Caissier", i, s.Cashier);
                dt.SetValue("Ouverture", i, s.OpenBal);
                dt.SetValue("Theorique", i, s.TheoBal);
                dt.SetValue("Compte", i, s.CountBal);
                dt.SetValue("Ecart", i, s.Diff);
                dt.SetValue("Statut", i, s.Status == SessionStatus.Open ? "Ouverte" : "Clôturée");
            }

            ((Matrix)_form.Items.Item(FormIds.MatSessions).Specific).LoadFromDataSource();
            ClearTransMatrix();
        }

        private void LoadSessionDetail(SessionRow session)
        {
            var lines = CashSessionService.GetTransactions(session.Code);

            DataTable dt = _form.DataSources.DataTables.Item("dtSessTrans");
            for (int i = dt.Rows.Count - 1; i >= 0; i--)
                dt.Rows.Remove(i);

            dt.Rows.Add(lines.Count);
            for (int i = 0; i < lines.Count; i++)
            {
                var t = lines[i];
                dt.SetValue("Heure", i, t.Time);
                dt.SetValue("Sens", i, t.Direction == Direction.Recette ? "Recette" : "Dépense");
                dt.SetValue("Type", i, t.TransTypeCode);
                dt.SetValue("Montant", i, t.Amount);
                dt.SetValue("Tiers", i, t.CardCode);
                dt.SetValue("Description", i, t.Description);
                dt.SetValue("Ecriture", i, t.JeDocEntry);
            }

            ((Matrix)_form.Items.Item(FormIds.MatSessionTrans).Specific).LoadFromDataSource();
        }

        private void ClearTransMatrix()
        {
            DataTable dt = _form.DataSources.DataTables.Item("dtSessTrans");
            for (int i = dt.Rows.Count - 1; i >= 0; i--)
                dt.Rows.Remove(i);
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

        private void AddLabel(string id, string caption, int left, int top, int width)
        {
            Item item = _form.Items.Add(id, BoFormItemTypes.it_STATIC);
            item.Left = left; item.Top = top; item.Width = width; item.Height = 15;
            ((StaticText)item.Specific).Caption = caption;
        }

        private EditText AddEdit(string id, int left, int top, int width)
        {
            Item item = _form.Items.Add(id, BoFormItemTypes.it_EDIT);
            item.Left = left; item.Top = top; item.Width = width; item.Height = 19;
            return (EditText)item.Specific;
        }

        private ComboBox AddCombo(string id, int left, int top, int width)
        {
            Item item = _form.Items.Add(id, BoFormItemTypes.it_COMBO_BOX);
            item.Left = left; item.Top = top; item.Width = width; item.Height = 19;
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
