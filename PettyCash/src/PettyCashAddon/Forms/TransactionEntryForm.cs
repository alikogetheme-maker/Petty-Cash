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
    /// L'instance est unique (créée par l'écran Session) et ne s'abonne
    /// qu'une fois aux événements.
    /// </summary>
    internal class TransactionEntryForm
    {
        private const string UdDir = "udDir";
        private const string UdType = "udType";
        private const string UdAmount = "udAmt";
        private const string UdCard = "udCard";
        private const string UdDescr = "udDescr";

        private readonly Application _app;
        private Action<Direction, TransactionType, double, string, string> _onConfirmed;
        private Action _onSaved;
        private Form _form;
        private Dictionary<string, TransactionType> _typesByCode = new Dictionary<string, TransactionType>();

        public TransactionEntryForm(Application app)
        {
            _app = app;
            _app.ItemEvent += App_ItemEvent;
        }

        /// <param name="onConfirmed">Enregistre la transaction ; une exception laisse la fenêtre ouverte.</param>
        /// <param name="onSaved">Appelé après fermeture de la fenêtre (ex. rafraîchir l'écran appelant) :
        /// on ne touche pas aux éléments d'un autre formulaire tant que celui-ci est actif.</param>
        public void Show(Action<Direction, TransactionType, double, string, string> onConfirmed, Action onSaved)
        {
            _onConfirmed = onConfirmed;
            _onSaved = onSaved;

            if (_form != null)
            {
                _form.Select();
                return;
            }

            FormCreationParams p = (FormCreationParams)_app.CreateObject(BoCreatableObjectType.cot_FormCreationParams);
            p.FormType = FormIds.TransactionEntryForm;
            p.UniqueID = FormIds.TransactionEntryForm;
            p.BorderStyle = BoFormBorderStyle.fbs_Fixed;

            _form = _app.Forms.AddEx(p);
            _form.Freeze(true);
            try
            {
                _form.Title = "Nouvelle transaction de caisse";
                _form.Left = 400;
                _form.Top = 200;
                _form.Width = 380;
                _form.Height = 260;

                UserDataSources uds = _form.DataSources.UserDataSources;
                uds.Add(UdDir, BoDataType.dt_SHORT_TEXT, 1);
                uds.Add(UdType, BoDataType.dt_SHORT_TEXT, 50);
                uds.Add(UdAmount, BoDataType.dt_SUM);
                uds.Add(UdCard, BoDataType.dt_SHORT_TEXT, 15);
                uds.Add(UdDescr, BoDataType.dt_SHORT_TEXT, 100);

                AddLabel(FormIds.LblDir, "Sens", 10, 15, FormIds.CmbDir);
                ComboBox cmbDir = AddCombo(FormIds.CmbDir, 130, 12, 200, UdDir);
                cmbDir.ValidValues.Add("R", "Recette (entrée d'argent)");
                cmbDir.ValidValues.Add("D", "Dépense (sortie d'argent)");
                uds.Item(UdDir).ValueEx = "D";

                AddLabel(FormIds.LblType, "Type d'opération", 10, 45, FormIds.CmbType);
                AddCombo(FormIds.CmbType, 130, 42, 200, UdType);

                AddLabel(FormIds.LblAmount, "Montant", 10, 75, FormIds.TxtAmount);
                AddEdit(FormIds.TxtAmount, 130, 72, 200, UdAmount);

                AddLabel(FormIds.LblCard, "Tiers (optionnel)", 10, 105, FormIds.TxtCard);
                AddEdit(FormIds.TxtCard, 130, 102, 200, UdCard);

                AddLabel(FormIds.LblDescr, "Description", 10, 135, FormIds.TxtDescr);
                AddEdit(FormIds.TxtDescr, 130, 132, 200, UdDescr);

                AddButton(FormIds.BtnOk, "Valider", 130, 180, 90);
                AddButton(FormIds.BtnCancel, "Annuler", 230, 180, 90);

                PopulateTypeCombo(Direction.Depense);
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
            _form.ActiveItem = FormIds.TxtAmount;
        }

        private void PopulateTypeCombo(Direction direction)
        {
            ComboBox cmb = (ComboBox)_form.Items.Item(FormIds.CmbType).Specific;

            _form.DataSources.UserDataSources.Item(UdType).ValueEx = "";
            while (cmb.ValidValues.Count > 0)
                cmb.ValidValues.Remove(0, BoSearchKey.psk_Index);

            _typesByCode = new Dictionary<string, TransactionType>();
            foreach (var t in CashSessionService.GetTransactionTypes(direction))
            {
                cmb.ValidValues.Add(t.Code, t.Name);
                _typesByCode[t.Code] = t;
            }

            if (_typesByCode.Count == 0)
                _app.StatusBar.SetText("Aucun type d'opération paramétré pour ce sens (table @PC_TTYPE).", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
        }

        private void App_ItemEvent(string formUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            if (formUID != FormIds.TransactionEntryForm)
                return;

            if (pVal.EventType == BoEventTypes.et_FORM_UNLOAD && !pVal.BeforeAction)
            {
                _form = null;
                return;
            }

            if (pVal.BeforeAction)
                return;

            try
            {
                if (pVal.EventType == BoEventTypes.et_COMBO_SELECT && pVal.ItemUID == FormIds.CmbDir)
                {
                    PopulateTypeCombo(SelectedDirection());
                    return;
                }

                if (pVal.EventType != BoEventTypes.et_ITEM_PRESSED || !pVal.ActionSuccess)
                    return;

                if (pVal.ItemUID == FormIds.BtnCancel)
                    _form.Close();
                else if (pVal.ItemUID == FormIds.BtnOk)
                    HandleOk();
            }
            catch (Exception ex)
            {
                // Le formulaire reste ouvert : l'utilisateur peut corriger et revalider
                _app.MessageBox(ex.Message);
            }
        }

        private Direction SelectedDirection()
        {
            return _form.DataSources.UserDataSources.Item(UdDir).ValueEx == "R" ? Direction.Recette : Direction.Depense;
        }

        private void HandleOk()
        {
            UserDataSources uds = _form.DataSources.UserDataSources;

            string typeCode = uds.Item(UdType).ValueEx;
            if (string.IsNullOrEmpty(typeCode) || !_typesByCode.TryGetValue(typeCode, out TransactionType type))
            {
                _app.StatusBar.SetText("Sélectionnez un type d'opération.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                return;
            }

            if (!double.TryParse(uds.Item(UdAmount).ValueEx, NumberStyles.Float, CultureInfo.InvariantCulture, out double amount) || amount <= 0)
            {
                _app.StatusBar.SetText("Montant invalide : saisissez un montant positif.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                return;
            }

            // Comptabilisation : en cas d'erreur, l'exception remonte au gestionnaire
            // d'événement qui l'affiche, et rien n'est enregistré (transaction DI API).
            _onConfirmed(SelectedDirection(), type, amount, uds.Item(UdCard).ValueEx, uds.Item(UdDescr).ValueEx);

            _form.Close();
            _app.StatusBar.SetText("Transaction enregistrée et comptabilisée.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

            // La transaction est enregistrée : une erreur de rafraîchissement ne doit
            // pas laisser croire le contraire (et inciter à valider une seconde fois).
            try
            {
                _onSaved?.Invoke();
            }
            catch (Exception ex)
            {
                _app.MessageBox("La transaction est bien enregistrée, mais l'écran de session n'a pas pu être actualisé : " +
                                ex.Message + "\n\nCliquez sur « Actualiser » dans l'écran de session.");
            }
        }

        // ---------- Helpers de construction d'UI ----------

        private void AddLabel(string id, string caption, int left, int top, string linkTo)
        {
            Item item = _form.Items.Add(id, BoFormItemTypes.it_STATIC);
            item.Left = left; item.Top = top; item.Width = 110; item.Height = 15;
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
