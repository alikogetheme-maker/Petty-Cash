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
    ///
    /// Les champs sont liés à des UserDataSources : c'est ce qui permet
    /// d'afficher une valeur dans un champ désactivé (affecter EditText.Value
    /// sur un champ désactivé lève "Form item not editable") et de laisser
    /// SAP gérer le format des montants selon le paramétrage de la société.
    /// </summary>
    internal class CashSessionFormController
    {
        private const string UdShift = "udShift";
        private const string UdCashier = "udCash";
        private const string UdStatus = "udStat";
        private const string UdOpenBal = "udOpen";
        private const string UdTheoBal = "udTheo";
        private const string UdCountBal = "udCount";
        private const string UdDiff = "udDiff";

        private readonly Application _app;
        private readonly TransactionEntryForm _entryForm;
        private Form _form;
        private SessionRow _session;

        public CashSessionFormController(Application app)
        {
            _app = app;
            _entryForm = new TransactionEntryForm(app);
            _app.ItemEvent += App_ItemEvent;
        }

        public void ShowOrActivate()
        {
            if (_form != null)
            {
                _form.Select();
                RefreshFromServer();
                return;
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
            _form.Freeze(true);
            try
            {
                _form.Title = "Session de caisse";
                _form.Width = 640;
                _form.Height = 430;

                UserDataSources uds = _form.DataSources.UserDataSources;
                uds.Add(UdShift, BoDataType.dt_SHORT_TEXT, 1);
                uds.Add(UdCashier, BoDataType.dt_SHORT_TEXT, 50);
                uds.Add(UdStatus, BoDataType.dt_SHORT_TEXT, 50);
                uds.Add(UdOpenBal, BoDataType.dt_SHORT_TEXT, 30);
                uds.Add(UdTheoBal, BoDataType.dt_SHORT_TEXT, 30);
                uds.Add(UdCountBal, BoDataType.dt_SHORT_TEXT, 30);
                uds.Add(UdDiff, BoDataType.dt_SHORT_TEXT, 30);

                AddLabel(FormIds.LblShift, "Quart", 10, 15, 90, FormIds.CmbShift);
                ComboBox cmbShift = AddCombo(FormIds.CmbShift, 110, 12, 150);
                cmbShift.DataBind.SetBound(true, "", UdShift);
                cmbShift.ValidValues.Add("M", "Matin");
                cmbShift.ValidValues.Add("A", "Après-midi");
                cmbShift.ValidValues.Add("S", "Soir");
                SetText(UdShift, "M");

                AddLabel(FormIds.LblCashier, "Caissier", 280, 15, 70, FormIds.TxtCashier);
                AddEdit(FormIds.TxtCashier, 360, 12, 260, UdCashier);

                AddLabel(FormIds.LblStatus, "Statut", 10, 45, 90, FormIds.TxtStatus);
                AddEdit(FormIds.TxtStatus, 110, 42, 260, UdStatus).Enabled = false;

                AddLabel(FormIds.LblOpenBal, "Solde ouverture", 10, 75, 95, FormIds.TxtOpenBal);
                AddEdit(FormIds.TxtOpenBal, 110, 72, 150, UdOpenBal).Enabled = false;

                AddLabel(FormIds.LblTheoBal, "Solde théorique", 280, 75, 95, FormIds.TxtTheoBal);
                AddEdit(FormIds.TxtTheoBal, 380, 72, 150, UdTheoBal).Enabled = false;

                AddLabel(FormIds.LblCountBal, "Solde compté", 10, 105, 95, FormIds.TxtCountBal);
                AddEdit(FormIds.TxtCountBal, 110, 102, 150, UdCountBal);

                AddLabel(FormIds.LblDiff, "Écart", 280, 105, 95, FormIds.TxtDiff);
                AddEdit(FormIds.TxtDiff, 380, 102, 150, UdDiff).Enabled = false;

                foreach (string amountItem in new[] { FormIds.TxtOpenBal, FormIds.TxtTheoBal, FormIds.TxtCountBal, FormIds.TxtDiff })
                    _form.Items.Item(amountItem).RightJustified = true;

                BuildMatrix();

                AddButton(FormIds.BtnOpen, "Ouvrir session", 10, 360, 130);
                AddButton(FormIds.BtnAddTrans, "Ajouter transaction", 150, 360, 150);
                AddButton(FormIds.BtnClose, "Clôturer la session", 310, 360, 150);
                AddButton(FormIds.BtnRefresh, "Actualiser", 520, 360, 100);
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
                RefreshFromServer();
            }
            catch (Exception ex)
            {
                _app.MessageBox(ex.Message);
            }
        }

        private void BuildMatrix()
        {
            DataTable dt = _form.DataSources.DataTables.Add("dtTrans");
            dt.Columns.Add("Heure", BoFieldsType.ft_AlphaNumeric, 8);
            dt.Columns.Add("Sens", BoFieldsType.ft_AlphaNumeric, 10);
            dt.Columns.Add("Type", BoFieldsType.ft_AlphaNumeric, 100);
            dt.Columns.Add("Montant", BoFieldsType.ft_Sum);
            dt.Columns.Add("Tiers", BoFieldsType.ft_AlphaNumeric, 15);
            dt.Columns.Add("Descr", BoFieldsType.ft_AlphaNumeric, 100);

            Item matItem = _form.Items.Add(FormIds.MatTrans, BoFormItemTypes.it_MATRIX);
            matItem.Left = 10; matItem.Top = 135; matItem.Width = 610; matItem.Height = 215;
            Matrix mat = (Matrix)matItem.Specific;
            mat.SelectionMode = BoMatrixSelect.ms_Single;

            Column num = mat.Columns.Add("#", BoFormItemTypes.it_EDIT);
            num.TitleObject.Caption = "#";
            num.Width = 20;
            num.Editable = false;

            AddMatrixColumn(mat, "colHeure", "Heure", 60, "Heure");
            AddMatrixColumn(mat, "colSens", "Sens", 60, "Sens");
            AddMatrixColumn(mat, "colType", "Type", 140, "Type");
            AddMatrixColumn(mat, "colMontant", "Montant", 90, "Montant");
            AddMatrixColumn(mat, "colTiers", "Tiers", 70, "Tiers");
            AddMatrixColumn(mat, "colDescr", "Description", 150, "Descr");
        }

        private void AddMatrixColumn(Matrix mat, string colId, string caption, int width, string boundColumn)
        {
            Column col = mat.Columns.Add(colId, BoFormItemTypes.it_EDIT);
            col.TitleObject.Caption = caption;
            col.Width = width;
            col.Editable = false;
            col.DataBind.Bind("dtTrans", boundColumn);
        }

        private void RefreshFromServer()
        {
            if (_form == null)
                return;

            _session = CashSessionService.GetOpenSession();
            bool hasOpenSession = _session != null;

            // SAP refuse de manipuler le focus ou les éléments d'un formulaire
            // qui n'est pas actif (erreur 66000-109) : on le ramène au premier plan.
            if (!_form.Selected)
                _form.Select();

            _form.Freeze(true);
            try
            {
                // On active d'abord ce qui doit l'être, on y place le focus, puis
                // on désactive le reste : SAP refuse de désactiver le champ actif.
                if (hasOpenSession)
                {
                    SetEnabled(FormIds.TxtCountBal, true);
                    _form.ActiveItem = FormIds.TxtCountBal;
                    SetEnabled(FormIds.CmbShift, false);
                    SetEnabled(FormIds.TxtCashier, false);
                }
                else
                {
                    SetEnabled(FormIds.CmbShift, true);
                    SetEnabled(FormIds.TxtCashier, true);
                    _form.ActiveItem = FormIds.TxtCashier;
                    SetEnabled(FormIds.TxtCountBal, false);
                }

                SetEnabled(FormIds.BtnOpen, !hasOpenSession);
                SetEnabled(FormIds.BtnAddTrans, hasOpenSession);
                SetEnabled(FormIds.BtnClose, hasOpenSession);

                if (hasOpenSession)
                {
                    SetText(UdShift, EnumCodes.ToCode(_session.Shift));
                    SetText(UdCashier, _session.Cashier);
                    SetText(UdStatus, "Ouverte le " + _session.CashDate.ToString("dd/MM/yyyy") + " (" + ShiftLabel(_session.Shift) + ")");
                    SetAmount(UdOpenBal, _session.OpenBal);
                    SetAmount(UdTheoBal, _session.TheoBal);
                    SetText(UdDiff, "");
                    RefreshMatrix();
                }
                else
                {
                    SetText(UdStatus, "Aucune session ouverte");
                    // Solde qui sera repris à l'ouverture, pour information
                    string forecastError = null;
                    try { SetAmount(UdOpenBal, CashSessionService.GetNextOpeningBalance()); }
                    catch (Exception ex) { SetText(UdOpenBal, ""); forecastError = ex.Message; }
                    SetText(UdTheoBal, "");
                    SetText(UdCountBal, "");
                    SetText(UdDiff, "");
                    ClearMatrix();

                    if (forecastError != null)
                        _app.StatusBar.SetText(forecastError, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
                }
            }
            finally
            {
                _form.Freeze(false);
            }
        }

        private void RefreshMatrix()
        {
            DataTable dt = _form.DataSources.DataTables.Item("dtTrans");
            dt.Rows.Clear();

            var lines = CashSessionService.GetTransactions(_session.Code);
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
            }

            ((Matrix)_form.Items.Item(FormIds.MatTrans).Specific).LoadFromDataSource();
        }

        private void ClearMatrix()
        {
            _form.DataSources.DataTables.Item("dtTrans").Rows.Clear();
            ((Matrix)_form.Items.Item(FormIds.MatTrans).Specific).LoadFromDataSource();
        }

        private void App_ItemEvent(string formUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            if (formUID != FormIds.SessionForm)
                return;

            if (pVal.EventType == BoEventTypes.et_FORM_UNLOAD && !pVal.BeforeAction)
            {
                _form = null;
                _session = null;
                return;
            }

            if (pVal.BeforeAction || pVal.EventType != BoEventTypes.et_ITEM_PRESSED || !pVal.ActionSuccess)
                return;

            try
            {
                if (pVal.ItemUID == FormIds.BtnOpen)
                    HandleOpen();
                else if (pVal.ItemUID == FormIds.BtnAddTrans)
                    HandleAddTrans();
                else if (pVal.ItemUID == FormIds.BtnClose)
                    HandleClose();
                else if (pVal.ItemUID == FormIds.BtnRefresh)
                    RefreshFromServer();
            }
            catch (Exception ex)
            {
                _app.MessageBox(ex.Message);
            }
        }

        private void HandleOpen()
        {
            string shiftCode = GetText(UdShift).Trim();
            string cashier = GetText(UdCashier).Trim();

            if (shiftCode.Length == 0)
            {
                _app.MessageBox("Sélectionnez le quart.");
                return;
            }
            if (cashier.Length == 0)
            {
                _app.MessageBox("Renseignez le nom du caissier.");
                return;
            }

            Shift shift = EnumCodes.ShiftFromCode(shiftCode);
            double opening = CashSessionService.GetNextOpeningBalance();

            int answer = _app.MessageBox(
                "Ouvrir la session " + ShiftLabel(shift) + " pour " + cashier + " ?\n\n" +
                "Solde d'ouverture repris : " + CashSessionService.FormatAmount(opening) + "\n" +
                "Vérifiez que ce montant correspond à l'argent présent dans la caisse.", 1, "Oui", "Non");
            if (answer != 1)
                return;

            _session = CashSessionService.OpenSession(shift, cashier);
            RefreshFromServer();
            _app.StatusBar.SetText("Session ouverte, solde d'ouverture " + CashSessionService.FormatAmount(_session.OpenBal), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
        }

        private void HandleAddTrans()
        {
            if (_session == null)
                return;

            string sessionCode = _session.Code;
            _entryForm.Show(
                (direction, type, amount, cardCode, description) =>
                    CashSessionService.AddTransaction(sessionCode, direction, type, amount, cardCode, description),
                RefreshFromServer);
        }

        private void HandleClose()
        {
            if (_session == null)
                return;

            string raw = GetText(UdCountBal).Trim();
            if (!CashSessionService.TryParseAmount(raw, out double counted))
            {
                _app.MessageBox("Saisissez le solde compté (argent physiquement présent dans la caisse) avant de clôturer.");
                return;
            }

            // Solde théorique relu depuis la base au cas où un autre poste aurait saisi entre-temps
            RefreshFromServer();
            if (_session == null)
                return;
            SetAmount(UdCountBal, counted);

            double expectedDiff = counted - _session.TheoBal;
            int answer = _app.MessageBox(
                "Solde théorique : " + CashSessionService.FormatAmount(_session.TheoBal) + "\n" +
                "Solde compté : " + CashSessionService.FormatAmount(counted) + "\n" +
                "Écart : " + CashSessionService.FormatAmount(expectedDiff) + (expectedDiff > 0 ? " (excédent)" : expectedDiff < 0 ? " (manquant)" : "") + "\n\n" +
                "Confirmer la clôture de la session ? Elle ne pourra plus être modifiée.", 2, "Oui", "Non");

            if (answer != 1)
                return;

            SessionRow closed = CashSessionService.CloseSession(_session.Code, counted);
            RefreshFromServer();
            _app.MessageBox(
                "Session clôturée.\n\n" +
                "Solde théorique : " + CashSessionService.FormatAmount(closed.TheoBal) + "\n" +
                "Solde compté : " + CashSessionService.FormatAmount(closed.CountBal) + "\n" +
                "Écart comptabilisé : " + CashSessionService.FormatAmount(closed.Diff));
        }

        // ---------- Helpers ----------

        private static string ShiftLabel(Shift shift)
        {
            switch (shift)
            {
                case Shift.ApresMidi: return "Après-midi";
                case Shift.Soir: return "Soir";
                default: return "Matin";
            }
        }

        private void SetEnabled(string itemId, bool enabled)
        {
            Item item = _form.Items.Item(itemId);
            if (item.Enabled != enabled)
                item.Enabled = enabled;
        }

        private string GetText(string udsId)
        {
            return _form.DataSources.UserDataSources.Item(udsId).ValueEx ?? "";
        }

        private void SetText(string udsId, string text)
        {
            _form.DataSources.UserDataSources.Item(udsId).ValueEx = text ?? "";
        }

        private void SetAmount(string udsId, double amount)
        {
            _form.DataSources.UserDataSources.Item(udsId).ValueEx = CashSessionService.FormatAmount(amount);
        }

        private void AddLabel(string id, string caption, int left, int top, int width, string linkTo)
        {
            Item item = _form.Items.Add(id, BoFormItemTypes.it_STATIC);
            item.Left = left; item.Top = top; item.Width = width; item.Height = 15;
            item.LinkTo = linkTo;
            ((StaticText)item.Specific).Caption = caption;
        }

        private Item AddEdit(string id, int left, int top, int width, string udsId)
        {
            Item item = _form.Items.Add(id, BoFormItemTypes.it_EDIT);
            item.Left = left; item.Top = top; item.Width = width; item.Height = 19;
            ((EditText)item.Specific).DataBind.SetBound(true, "", udsId);
            return item;
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
