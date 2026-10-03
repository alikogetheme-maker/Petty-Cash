using System;
using System.Linq;
using SAPbouiCOM;
using PettyCashAddon.Models;
using PettyCashAddon.Services;

namespace PettyCashAddon.Forms
{
    /// <summary>
    /// Écran principal "Session de caisse" de l'utilisateur SAP connecté :
    /// choix d'une caisse libre, ouverture avec billetage, saisie des
    /// transactions, clôture avec billetage.
    ///
    /// Les champs sont liés à des UserDataSources : c'est ce qui permet
    /// d'afficher une valeur dans un champ désactivé (affecter EditText.Value
    /// sur un champ désactivé lève "Form item not editable").
    /// </summary>
    internal class CashSessionFormController
    {
        private const string UdBox = "udBox";
        private const string UdShift = "udShift";
        private const string UdUser = "udUser";
        private const string UdStatus = "udStat";
        private const string UdOpenBal = "udOpen";
        private const string UdTheoBal = "udTheo";
        private const string UdCountBal = "udCount";
        private const string UdDiff = "udDiff";
        private const string UdSink = "udSink";
        // Petit champ invisible à l'œil qui reçoit le focus : SAP refuse de
        // désactiver l'élément actif, et ni les boutons ni la grille ne le prennent.
        private const string FocusSink = "txtSink";

        private readonly Application _app;
        private readonly TransactionEntryForm _entryForm;
        private readonly BillCountForm _billForm;
        private Form _form;
        private SessionRow _session;

        public CashSessionFormController(Application app, BillCountForm billForm)
        {
            _app = app;
            _billForm = billForm;
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

        /// <summary>Rafraîchit l'écran s'il est ouvert (appelé après une clôture depuis le rapport).</summary>
        public void RefreshIfOpen()
        {
            if (_form != null)
                RefreshFromServer();
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
                uds.Add(UdBox, BoDataType.dt_SHORT_TEXT, 20);
                uds.Add(UdShift, BoDataType.dt_SHORT_TEXT, 1);
                uds.Add(UdUser, BoDataType.dt_SHORT_TEXT, 100);
                uds.Add(UdStatus, BoDataType.dt_SHORT_TEXT, 100);
                uds.Add(UdOpenBal, BoDataType.dt_SHORT_TEXT, 30);
                uds.Add(UdTheoBal, BoDataType.dt_SHORT_TEXT, 30);
                uds.Add(UdCountBal, BoDataType.dt_SHORT_TEXT, 30);
                uds.Add(UdDiff, BoDataType.dt_SHORT_TEXT, 30);
                uds.Add(UdSink, BoDataType.dt_SHORT_TEXT, 1);

                AddLabel(FormIds.LblBox, "Caisse", 10, 15, 95, FormIds.CmbBox);
                AddCombo(FormIds.CmbBox, 110, 12, 200).DataBind.SetBound(true, "", UdBox);

                AddLabel(FormIds.LblShift, "Quart", 330, 15, 65, FormIds.CmbShift);
                ComboBox cmbShift = AddCombo(FormIds.CmbShift, 400, 12, 150);
                cmbShift.DataBind.SetBound(true, "", UdShift);
                cmbShift.ValidValues.Add("M", "Matin");
                cmbShift.ValidValues.Add("A", "Après-midi");
                cmbShift.ValidValues.Add("S", "Soir");
                SetText(UdShift, DefaultShiftCode());

                AddLabel(FormIds.LblUser, "Utilisateur", 10, 45, 95, FormIds.TxtUser);
                AddEdit(FormIds.TxtUser, 110, 42, 200, UdUser).Enabled = false;

                AddLabel(FormIds.LblStatus, "Statut", 330, 45, 65, FormIds.TxtStatus);
                AddEdit(FormIds.TxtStatus, 400, 42, 220, UdStatus).Enabled = false;

                AddLabel(FormIds.LblOpenBal, "Solde ouverture", 10, 75, 95, FormIds.TxtOpenBal);
                AddEdit(FormIds.TxtOpenBal, 110, 72, 150, UdOpenBal).Enabled = false;

                AddLabel(FormIds.LblTheoBal, "Solde théorique", 330, 75, 65, FormIds.TxtTheoBal);
                AddEdit(FormIds.TxtTheoBal, 400, 72, 150, UdTheoBal).Enabled = false;

                AddLabel(FormIds.LblCountBal, "Solde compté", 10, 105, 95, FormIds.TxtCountBal);
                AddEdit(FormIds.TxtCountBal, 110, 102, 150, UdCountBal).Enabled = false;

                AddLabel(FormIds.LblDiff, "Écart", 330, 105, 65, FormIds.TxtDiff);
                AddEdit(FormIds.TxtDiff, 400, 102, 150, UdDiff).Enabled = false;

                foreach (string amountItem in new[] { FormIds.TxtOpenBal, FormIds.TxtTheoBal, FormIds.TxtCountBal, FormIds.TxtDiff })
                    _form.Items.Item(amountItem).RightJustified = true;

                BuildMatrix();

                AddButton(FormIds.BtnOpen, "Ouvrir session", 10, 360, 130);
                AddButton(FormIds.BtnAddTrans, "Ajouter transaction", 150, 360, 150);
                AddButton(FormIds.BtnClose, "Clôturer la session", 310, 360, 150);
                AddButton(FormIds.BtnRefresh, "Actualiser", 520, 360, 100);

                AddEdit(FocusSink, 636, 392, 1, UdSink);
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

            SapUser user = CashSessionService.GetCurrentUser();
            _session = CashSessionService.GetOpenSessionForUser(user.Code);
            bool hasOpenSession = _session != null;

            // SAP refuse de manipuler le focus ou les éléments d'un formulaire
            // qui n'est pas actif (erreur 66000-109) : on le ramène au premier plan.
            if (!_form.Selected)
                _form.Select();

            _form.Freeze(true);
            try
            {
                // Focus sur le champ "puits" avant toute (dés)activation
                _form.ActiveItem = FocusSink;

                SetText(UdUser, user.Name + (user.IsSuperUser ? " (superutilisateur)" : ""));
                ComboBox cmbBox = (ComboBox)_form.Items.Item(FormIds.CmbBox).Specific;
                SetText(UdBox, "");
                while (cmbBox.ValidValues.Count > 0)
                    cmbBox.ValidValues.Remove(0, BoSearchKey.psk_Index);

                bool canOpen = false;
                if (hasOpenSession)
                {
                    cmbBox.ValidValues.Add(_session.CashBox, CashSessionService.CashBoxName(_session.CashBox));
                    SetText(UdBox, _session.CashBox);
                    SetText(UdShift, EnumCodes.ToCode(_session.Shift));
                    SetText(UdStatus, "Ouverte le " + _session.CashDate.ToString("dd/MM/yyyy") + " (" + ShiftLabel(_session.Shift) + ")");
                    SetAmount(UdOpenBal, _session.OpenBal);
                    SetAmount(UdTheoBal, _session.TheoBal);
                    SetText(UdCountBal, "");
                    SetText(UdDiff, "");
                    RefreshMatrix();
                }
                else
                {
                    var free = CashSessionService.GetCashBoxes(true).Where(b => b.IsFree).ToList();
                    foreach (var box in free)
                        cmbBox.ValidValues.Add(box.Code, box.Name);
                    canOpen = free.Count > 0;
                    if (canOpen)
                        SetText(UdBox, free[0].Code);

                    SetText(UdStatus, canOpen ? "Aucune session ouverte" : "Aucune caisse libre");
                    SetText(UdTheoBal, "");
                    SetText(UdCountBal, "");
                    SetText(UdDiff, "");
                    ShowOpeningForecast();
                    ClearMatrix();
                }

                SetEnabled(FormIds.CmbBox, !hasOpenSession && canOpen);
                SetEnabled(FormIds.CmbShift, !hasOpenSession);
                SetEnabled(FormIds.BtnOpen, !hasOpenSession && canOpen);
                SetEnabled(FormIds.BtnAddTrans, hasOpenSession);
                SetEnabled(FormIds.BtnClose, hasOpenSession);
            }
            finally
            {
                _form.Freeze(false);
            }
        }

        /// <summary>Affiche le solde qui sera attendu à l'ouverture de la caisse choisie.</summary>
        private void ShowOpeningForecast()
        {
            string box = GetText(UdBox);
            if (box.Length == 0)
            {
                SetText(UdOpenBal, "");
                return;
            }
            try
            {
                SetAmount(UdOpenBal, CashSessionService.GetNextOpeningBalance(box));
            }
            catch (Exception ex)
            {
                SetText(UdOpenBal, "");
                _app.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
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

            if (pVal.BeforeAction)
                return;

            try
            {
                if (pVal.EventType == BoEventTypes.et_COMBO_SELECT && pVal.ItemUID == FormIds.CmbBox && _session == null)
                {
                    ShowOpeningForecast();
                    return;
                }

                if (pVal.EventType != BoEventTypes.et_ITEM_PRESSED || !pVal.ActionSuccess)
                    return;

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
            string boxCode = GetText(UdBox).Trim();
            string shiftCode = GetText(UdShift).Trim();
            if (boxCode.Length == 0)
            {
                _app.MessageBox("Choisissez une caisse.");
                return;
            }
            if (shiftCode.Length == 0)
            {
                _app.MessageBox("Sélectionnez le quart.");
                return;
            }

            Shift shift = EnumCodes.ShiftFromCode(shiftCode);
            string boxName = CashSessionService.CashBoxName(boxCode);
            double expected = CashSessionService.GetNextOpeningBalance(boxCode);
            SessionRow opened = null;

            _billForm.Show(
                "Billetage d'ouverture",
                boxName + " — " + ShiftLabel(shift) + " — " + CashSessionService.GetCurrentUser().Name,
                expected,
                "Confirmer l'ouverture de la session ?",
                lines => opened = CashSessionService.OpenSession(boxCode, shift, lines),
                () =>
                {
                    RefreshFromServer();
                    string msg = "Session ouverte sur « " + boxName + " », solde d'ouverture " + CashSessionService.FormatAmount(opened.OpenBal);
                    if (opened.OpenDiff != 0)
                        msg += " (écart d'ouverture " + CashSessionService.FormatAmount(opened.OpenDiff) + " comptabilisé)";
                    _app.StatusBar.SetText(msg, BoMessageTime.bmt_Medium, opened.OpenDiff == 0 ? BoStatusBarMessageType.smt_Success : BoStatusBarMessageType.smt_Warning);
                });
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

            // Solde théorique relu depuis la base avant le billetage
            RefreshFromServer();
            if (_session == null)
                return;

            ShowCloseCount(_app, _billForm, _session, RefreshFromServer);
        }

        /// <summary>
        /// Lance le billetage de clôture d'une session (aussi utilisé par le
        /// rapport pour la clôture par un superutilisateur).
        /// </summary>
        internal static void ShowCloseCount(Application app, BillCountForm billForm, SessionRow session, Action afterClose)
        {
            if (!CashSessionService.CanClose(session))
                throw new InvalidOperationException("Seul l'utilisateur qui a ouvert cette session (" + session.Cashier + ") ou un superutilisateur peut la clôturer.");

            SessionRow closed = null;
            billForm.Show(
                "Billetage de clôture",
                CashSessionService.CashBoxName(session.CashBox) + " — " + ShiftLabel(session.Shift) + " — ouverte par " + session.Cashier,
                session.TheoBal,
                "Confirmer la clôture de la session ? Elle ne pourra plus être modifiée.",
                lines => closed = CashSessionService.CloseSession(session.Code, lines),
                () =>
                {
                    afterClose?.Invoke();
                    app.MessageBox(
                        "Session clôturée.\n\n" +
                        "Solde théorique : " + CashSessionService.FormatAmount(closed.TheoBal) + "\n" +
                        "Solde compté : " + CashSessionService.FormatAmount(closed.CountBal) + "\n" +
                        "Écart comptabilisé : " + CashSessionService.FormatAmount(closed.Diff));
                });
        }

        // ---------- Helpers ----------

        internal static string ShiftLabel(Shift shift)
        {
            switch (shift)
            {
                case Shift.ApresMidi: return "Après-midi";
                case Shift.Soir: return "Soir";
                default: return "Matin";
            }
        }

        /// <summary>Quart proposé selon l'heure : matin avant 12 h, après-midi avant 18 h, soir ensuite.</summary>
        private static string DefaultShiftCode()
        {
            int h = DateTime.Now.Hour;
            return h < 12 ? "M" : h < 18 ? "A" : "S";
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
