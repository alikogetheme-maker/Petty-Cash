using System;
using System.Globalization;
using SAPbouiCOM;
using PettyCashAddon.Models;
using PettyCashAddon.Services;

namespace PettyCashAddon.Forms
{
    /// <summary>
    /// Écran principal "Session de caisse" : ouverture, saisie des
    /// transactions du quart en cours, clôture avec comptage physique.
    /// </summary>
    internal class CashSessionFormController
    {
        private readonly Application _app;
        private Form _form;
        private SessionRow _session;

        public CashSessionFormController(Application app)
        {
            _app = app;
        }

        public void ShowOrActivate()
        {
            foreach (Form f in _app.Forms)
            {
                if (f.TypeEx == FormIds.SessionForm)
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
            p.FormType = FormIds.SessionForm;
            p.UniqueID = FormIds.SessionForm;
            p.BorderStyle = BoFormBorderStyle.fbs_Fixed;

            _form = _app.Forms.AddEx(p);
            _form.Title = "Session de caisse";
            _form.Width = 640;
            _form.Height = 430;

            AddLabel(FormIds.LblShift, "Quart", 10, 15, 90);
            ComboBox cmbShift = AddCombo(FormIds.CmbShift, 110, 12, 150);
            cmbShift.ValidValues.Add("M", "Matin");
            cmbShift.ValidValues.Add("A", "Après-midi");
            cmbShift.ValidValues.Add("S", "Soir");
            cmbShift.Select("M", BoSearchKey.psk_ByValue);

            AddLabel(FormIds.LblCashier, "Caissier", 280, 15, 70);
            AddEdit(FormIds.TxtCashier, 360, 12, 260);

            AddLabel(FormIds.LblStatus, "Statut", 10, 45, 90);
            AddEdit(FormIds.TxtStatus, 110, 42, 150).Item.Enabled = false;

            AddLabel(FormIds.LblOpenBal, "Solde ouverture", 10, 75, 90);
            AddEdit(FormIds.TxtOpenBal, 110, 72, 150).Item.Enabled = false;

            AddLabel(FormIds.LblTheoBal, "Solde théorique", 280, 75, 90);
            AddEdit(FormIds.TxtTheoBal, 380, 72, 150).Item.Enabled = false;

            AddLabel(FormIds.LblCountBal, "Solde compté", 10, 105, 90);
            AddEdit(FormIds.TxtCountBal, 110, 102, 150);

            AddLabel(FormIds.LblDiff, "Écart", 280, 105, 90);
            AddEdit(FormIds.TxtDiff, 380, 102, 150).Item.Enabled = false;

            BuildMatrix();

            AddButton(FormIds.BtnOpen, "Ouvrir session", 10, 360, 130);
            AddButton(FormIds.BtnAddTrans, "Ajouter transaction", 150, 360, 150);
            AddButton(FormIds.BtnClose, "Clôturer la session", 310, 360, 150);

            _form.Visible = true;

            _app.ItemEvent += App_ItemEvent;

            RefreshFromServer();
        }

        private void BuildMatrix()
        {
            DataTable dt = _form.DataSources.DataTables.Add("dtTrans");
            dt.Columns.Add("Heure", BoFieldTypes.db_Alpha);
            dt.Columns.Add("Sens", BoFieldTypes.db_Alpha);
            dt.Columns.Add("Type", BoFieldTypes.db_Alpha);
            dt.Columns.Add("Montant", BoFieldTypes.db_Float);
            dt.Columns.Add("Tiers", BoFieldTypes.db_Alpha);
            dt.Columns.Add("Description", BoFieldTypes.db_Alpha);

            Item matItem = _form.Items.Add(FormIds.MatTrans, BoFormItemTypes.it_MATRIX);
            matItem.Left = 10; matItem.Top = 135; matItem.Width = 610; matItem.Height = 215;
            Matrix mat = (Matrix)matItem.Specific;

            AddMatrixColumn(mat, "colHeure", "Heure", 60, "Heure");
            AddMatrixColumn(mat, "colSens", "Sens", 60, "Sens");
            AddMatrixColumn(mat, "colType", "Type", 150, "Type");
            AddMatrixColumn(mat, "colMontant", "Montant", 90, "Montant");
            AddMatrixColumn(mat, "colTiers", "Tiers", 80, "Tiers");
            AddMatrixColumn(mat, "colDescr", "Description", 160, "Description");
        }

        private void AddMatrixColumn(Matrix mat, string colId, string caption, int width, string boundColumn)
        {
            Column col = mat.Columns.Add(colId, BoFormItemTypes.it_EDIT);
            col.Title = caption;
            col.Width = width;
            col.Editable = false;
            col.DataBind.SetBound(true, "dtTrans", boundColumn);
        }

        private void RefreshFromServer()
        {
            _session = CashSessionService.GetOpenSession();

            bool hasOpenSession = _session != null;

            _form.Items.Item(FormIds.CmbShift).Enabled = !hasOpenSession;
            _form.Items.Item(FormIds.TxtCashier).Enabled = !hasOpenSession;
            _form.Items.Item(FormIds.BtnOpen).Enabled = !hasOpenSession;
            _form.Items.Item(FormIds.BtnAddTrans).Enabled = hasOpenSession;
            _form.Items.Item(FormIds.BtnClose).Enabled = hasOpenSession;
            _form.Items.Item(FormIds.TxtCountBal).Enabled = hasOpenSession;

            if (hasOpenSession)
            {
                SetText(FormIds.TxtStatus, "Ouverte");
                SetText(FormIds.TxtOpenBal, _session.OpenBal.ToString("N2", CultureInfo.CurrentCulture));
                SetText(FormIds.TxtTheoBal, _session.TheoBal.ToString("N2", CultureInfo.CurrentCulture));
                SetText(FormIds.TxtDiff, "");
                RefreshMatrix();
            }
            else
            {
                SetText(FormIds.TxtStatus, "Aucune session ouverte");
                SetText(FormIds.TxtOpenBal, "");
                SetText(FormIds.TxtTheoBal, "");
                SetText(FormIds.TxtCountBal, "");
                SetText(FormIds.TxtDiff, "");
                ClearMatrix();
            }
        }

        private void RefreshMatrix()
        {
            ClearMatrix();
            if (_session == null)
                return;

            DataTable dt = _form.DataSources.DataTables.Item("dtTrans");
            var lines = CashSessionService.GetTransactions(_session.Code);
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
            }

            Matrix mat = (Matrix)_form.Items.Item(FormIds.MatTrans).Specific;
            mat.LoadFromDataSource();
        }

        private void ClearMatrix()
        {
            DataTable dt = _form.DataSources.DataTables.Item("dtTrans");
            for (int i = dt.Rows.Count - 1; i >= 0; i--)
                dt.Rows.Remove(i);
            Matrix mat = (Matrix)_form.Items.Item(FormIds.MatTrans).Specific;
            mat.LoadFromDataSource();
        }

        private void App_ItemEvent(string formUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            if (formUID != FormIds.SessionForm)
                return;

            if (pVal.EventType == BoEventTypes.et_FORM_CLOSE)
            {
                _app.ItemEvent -= App_ItemEvent;
                return;
            }

            if (pVal.BeforeAction || pVal.EventType != BoEventTypes.et_CLICK)
                return;

            try
            {
                if (pVal.ItemUID == FormIds.BtnOpen)
                    HandleOpen();
                else if (pVal.ItemUID == FormIds.BtnAddTrans)
                    HandleAddTrans();
                else if (pVal.ItemUID == FormIds.BtnClose)
                    HandleClose();
            }
            catch (Exception ex)
            {
                _app.MessageBox(ex.Message);
            }
        }

        private void HandleOpen()
        {
            ComboBox cmbShift = (ComboBox)_form.Items.Item(FormIds.CmbShift).Specific;
            EditText txtCashier = (EditText)_form.Items.Item(FormIds.TxtCashier).Specific;

            Shift shift = cmbShift.Selected.Value == "A" ? Shift.ApresMidi
                        : cmbShift.Selected.Value == "S" ? Shift.Soir
                        : Shift.Matin;

            if (string.IsNullOrWhiteSpace(txtCashier.Value))
            {
                _app.MessageBox("Renseignez le nom du caissier.");
                return;
            }

            _session = CashSessionService.OpenSession(shift, txtCashier.Value, 0);
            RefreshFromServer();
            _app.StatusBar.SetText("Session ouverte, solde d'ouverture " + _session.OpenBal.ToString("N2"), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
        }

        private void HandleAddTrans()
        {
            if (_session == null)
                return;

            var popup = new TransactionEntryForm(_app, (direction, type, amount, cardCode, description) =>
            {
                CashSessionService.AddTransaction(_session, direction, type, amount, cardCode, description);
                _session = CashSessionService.LoadSession(_session.Code);
                RefreshFromServer();
            });
            popup.Show();
        }

        private void HandleClose()
        {
            if (_session == null)
                return;

            EditText txtCount = (EditText)_form.Items.Item(FormIds.TxtCountBal).Specific;
            if (!double.TryParse(txtCount.Value, NumberStyles.Any, CultureInfo.CurrentCulture, out double counted))
            {
                _app.MessageBox("Saisissez le solde compté avant de clôturer.");
                return;
            }

            double expectedDiff = counted - _session.TheoBal;
            int answer = _app.MessageBox(
                "Solde théorique : " + _session.TheoBal.ToString("N2") + "\n" +
                "Solde compté : " + counted.ToString("N2") + "\n" +
                "Écart : " + expectedDiff.ToString("N2") + "\n\n" +
                "Confirmer la clôture de la session ?", 1, "Oui", "Non");

            if (answer != 1)
                return;

            _session = CashSessionService.CloseSession(_session, counted);
            SetText(FormIds.TxtDiff, _session.Diff.ToString("N2", CultureInfo.CurrentCulture));
            _app.StatusBar.SetText("Session clôturée. Écart : " + _session.Diff.ToString("N2"), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            RefreshFromServer();
        }

        // ---------- Helpers ----------

        private void SetText(string itemId, string text)
        {
            ((EditText)_form.Items.Item(itemId).Specific).Value = text;
        }

        private void AddLabel(string id, string caption, int left, int top, int width)
        {
            Item item = _form.Items.Add(id, BoFormItemTypes.it_STATIC);
            item.Left = left; item.Top = top; item.Width = width; item.Height = 15;
            ((StaticText)item.Specific).Caption = caption;
        }

        private EditTextRef AddEdit(string id, int left, int top, int width)
        {
            Item item = _form.Items.Add(id, BoFormItemTypes.it_EDIT);
            item.Left = left; item.Top = top; item.Width = width; item.Height = 19;
            return new EditTextRef(item);
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

        /// <summary>Petit wrapper pour pouvoir chaîner ".Item.Enabled = false" juste après la création d'un champ.</summary>
        private struct EditTextRef
        {
            public readonly Item Item;
            public EditTextRef(Item item) { Item = item; }
        }
    }
}
