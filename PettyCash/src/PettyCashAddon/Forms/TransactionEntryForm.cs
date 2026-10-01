using System;
using System.Collections.Generic;
using System.Globalization;
using SAPbouiCOM;
using PettyCashAddon.Models;
using PettyCashAddon.Services;

namespace PettyCashAddon.Forms
{
    /// <summary>
    /// Popup de saisie d'une transaction de caisse. Le formulaire UI API
    /// n'étant pas modal au sens WinForms, on gère la validation via un
    /// callback déclenché quand l'utilisateur clique sur "Valider".
    /// </summary>
    internal class TransactionEntryForm
    {
        private readonly Application _app;
        private readonly Action<Direction, TransactionType, double, string, string> _onConfirmed;
        private Form _form;
        private Dictionary<string, TransactionType> _typesByCode;

        public TransactionEntryForm(Application app, Action<Direction, TransactionType, double, string, string> onConfirmed)
        {
            _app = app;
            _onConfirmed = onConfirmed;
        }

        public void Show()
        {
            FormCreationParams p = (FormCreationParams)_app.CreateObject(BoCreatableObjectType.cot_FormCreationParams);
            p.FormType = FormIds.TransactionEntryForm;
            p.UniqueID = FormIds.TransactionEntryForm;
            p.BorderStyle = BoFormBorderStyle.fbs_Fixed;

            _form = _app.Forms.AddEx(p);
            _form.Title = "Nouvelle transaction de caisse";
            _form.Left = 400;
            _form.Top = 200;
            _form.Width = 380;
            _form.Height = 260;

            AddLabel(FormIds.LblDir, "Sens", 10, 15);
            ComboBox cmbDir = AddCombo(FormIds.CmbDir, 130, 12, 200);
            cmbDir.ValidValues.Add("R", "Recette");
            cmbDir.ValidValues.Add("D", "Dépense");
            cmbDir.Select("R", BoSearchKey.psk_ByValue);

            AddLabel(FormIds.LblType, "Type d'opération", 10, 45);
            AddCombo(FormIds.CmbType, 130, 42, 200);
            PopulateTypeCombo(Direction.Recette);

            AddLabel(FormIds.LblAmount, "Montant", 10, 75);
            AddEdit(FormIds.TxtAmount, 130, 72, 200);

            AddLabel(FormIds.LblCard, "Tiers (optionnel)", 10, 105);
            AddEdit(FormIds.TxtCard, 130, 102, 200);

            AddLabel(FormIds.LblDescr, "Description", 10, 135);
            AddEdit(FormIds.TxtDescr, 130, 132, 200);

            AddButton(FormIds.BtnOk, "Valider", 130, 180, 90);
            AddButton(FormIds.BtnCancel, "Annuler", 230, 180, 90);

            _form.Visible = true;
            _app.ItemEvent += App_ItemEvent;
        }

        private void PopulateTypeCombo(Direction direction)
        {
            Item item = _form.Items.Item(FormIds.CmbType);
            ComboBox cmb = (ComboBox)item.Specific;

            while (cmb.ValidValues.Count > 0)
                cmb.ValidValues.Remove(0, BoSearchKey.psk_Index);

            _typesByCode = new Dictionary<string, TransactionType>();
            foreach (var t in CashSessionService.GetTransactionTypes(direction))
            {
                cmb.ValidValues.Add(t.Code, t.Name);
                _typesByCode[t.Code] = t;
            }
        }

        private void App_ItemEvent(string formUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            if (formUID != FormIds.TransactionEntryForm)
                return;

            if (pVal.EventType == BoEventTypes.et_FORM_CLOSE)
            {
                _app.ItemEvent -= App_ItemEvent;
                return;
            }

            if (pVal.BeforeAction)
                return;

            if (pVal.EventType == BoEventTypes.et_COMBO_SELECT && pVal.ItemUID == FormIds.CmbDir)
            {
                ComboBox cmbDir = (ComboBox)_form.Items.Item(FormIds.CmbDir).Specific;
                Direction dir = cmbDir.Selected.Value == "R" ? Direction.Recette : Direction.Depense;
                PopulateTypeCombo(dir);
                return;
            }

            if (pVal.EventType != BoEventTypes.et_CLICK)
                return;

            if (pVal.ItemUID == FormIds.BtnCancel)
            {
                _form.Close();
                return;
            }

            if (pVal.ItemUID == FormIds.BtnOk)
            {
                HandleOk();
            }
        }

        private void HandleOk()
        {
            ComboBox cmbDir = (ComboBox)_form.Items.Item(FormIds.CmbDir).Specific;
            ComboBox cmbType = (ComboBox)_form.Items.Item(FormIds.CmbType).Specific;
            EditText txtAmount = (EditText)_form.Items.Item(FormIds.TxtAmount).Specific;
            EditText txtCard = (EditText)_form.Items.Item(FormIds.TxtCard).Specific;
            EditText txtDescr = (EditText)_form.Items.Item(FormIds.TxtDescr).Specific;

            if (cmbType.Selected == null)
            {
                _app.StatusBar.SetText("Sélectionnez un type d'opération.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                return;
            }

            if (!double.TryParse(txtAmount.Value, NumberStyles.Any, CultureInfo.CurrentCulture, out double amount) || amount <= 0)
            {
                _app.StatusBar.SetText("Montant invalide.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                return;
            }

            Direction direction = cmbDir.Selected.Value == "R" ? Direction.Recette : Direction.Depense;
            TransactionType type = _typesByCode[cmbType.Selected.Value];

            _onConfirmed(direction, type, amount, txtCard.Value, txtDescr.Value);

            _app.ItemEvent -= App_ItemEvent;
            _form.Close();
        }

        // ---------- Helpers de construction d'UI ----------

        private void AddLabel(string id, string caption, int left, int top)
        {
            Item item = _form.Items.Add(id, BoFormItemTypes.it_STATIC);
            item.Left = left; item.Top = top; item.Width = 110; item.Height = 15;
            StaticText st = (StaticText)item.Specific;
            st.Caption = caption;
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
            Button btn = (Button)item.Specific;
            btn.Caption = caption;
        }
    }
}
