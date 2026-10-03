using System;
using System.Collections.Generic;
using System.Linq;
using SAPbouiCOM;
using PettyCashAddon.Services;

namespace PettyCashAddon.Forms
{
    /// <summary>
    /// Billetage : saisie du nombre de billets et de pièces par coupure, avec
    /// calcul du total compté et de l'écart par rapport au montant attendu.
    /// Utilisé à l'ouverture (contrôle du fonds repris) et à la clôture.
    /// Instance unique, abonnée une seule fois aux événements.
    /// </summary>
    internal class BillCountForm
    {
        private const string Dt = "dtBill";
        private const string UdInfo = "udInfo";
        private const string UdExp = "udExp";
        private const string UdTotal = "udTotal";
        private const string UdDiff = "udDiff";

        private readonly Application _app;
        private Form _form;
        private List<Denomination> _denoms = new List<Denomination>();
        private double _expected;
        private string _confirmQuestion;
        private Action<List<CountLine>> _onConfirmed;
        private Action _onSaved;

        public BillCountForm(Application app)
        {
            _app = app;
            _app.ItemEvent += App_ItemEvent;
        }

        public bool IsOpen => _form != null;

        /// <param name="title">Titre de la fenêtre (ex. "Billetage d'ouverture").</param>
        /// <param name="info">Ligne d'information (caisse, session).</param>
        /// <param name="expected">Montant attendu (solde repris ou solde théorique).</param>
        /// <param name="confirmQuestion">Question finale de la confirmation.</param>
        /// <param name="onConfirmed">Enregistre ; une exception laisse la fenêtre ouverte.</param>
        /// <param name="onSaved">Appelé après fermeture de la fenêtre.</param>
        public void Show(string title, string info, double expected, string confirmQuestion,
                         Action<List<CountLine>> onConfirmed, Action onSaved)
        {
            if (_form != null)
            {
                _form.Select();
                _app.StatusBar.SetText("Un billetage est déjà en cours : terminez-le ou annulez-le.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                return;
            }

            _denoms = CashSessionService.GetDenominations();
            if (_denoms.Count == 0)
                throw new InvalidOperationException("Aucune coupure active dans la table @" + Db.DenomTable + ".");

            _expected = expected;
            _confirmQuestion = confirmQuestion;
            _onConfirmed = onConfirmed;
            _onSaved = onSaved;

            FormCreationParams p = (FormCreationParams)_app.CreateObject(BoCreatableObjectType.cot_FormCreationParams);
            p.FormType = FormIds.BillForm;
            p.UniqueID = FormIds.BillForm;
            p.BorderStyle = BoFormBorderStyle.fbs_Fixed;

            _form = _app.Forms.AddEx(p);
            _form.Freeze(true);
            try
            {
                _form.Title = title;
                _form.Left = 380;
                _form.Top = 120;
                _form.Width = 450;
                _form.Height = 500;

                UserDataSources uds = _form.DataSources.UserDataSources;
                uds.Add(UdInfo, BoDataType.dt_SHORT_TEXT, 100);
                uds.Add(UdExp, BoDataType.dt_SHORT_TEXT, 30);
                uds.Add(UdTotal, BoDataType.dt_SHORT_TEXT, 30);
                uds.Add(UdDiff, BoDataType.dt_SHORT_TEXT, 30);

                Item infoItem = _form.Items.Add(FormIds.LblBillInfo, BoFormItemTypes.it_STATIC);
                infoItem.Left = 10; infoItem.Top = 10; infoItem.Width = 420; infoItem.Height = 15;
                ((StaticText)infoItem.Specific).Caption = info;

                AddLabel(FormIds.LblExpected, "Montant attendu", 10, 35, FormIds.TxtExpected);
                AddAmountField(FormIds.TxtExpected, 160, 32, UdExp);

                BuildMatrix();

                AddLabel(FormIds.LblTotal, "Total compté", 10, 375, FormIds.TxtTotal);
                AddAmountField(FormIds.TxtTotal, 160, 372, UdTotal);
                AddLabel(FormIds.LblBillDiff, "Écart", 10, 400, FormIds.TxtBillDiff);
                AddAmountField(FormIds.TxtBillDiff, 160, 397, UdDiff);

                AddButton(FormIds.BtnBillOk, "Valider", 10, 435, 90);
                AddButton(FormIds.BtnBillCancel, "Annuler", 110, 435, 90);

                uds.Item(UdExp).ValueEx = CashSessionService.FormatAmount(expected);
                UpdateTotals();
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
                // Curseur directement sur la première quantité
                Matrix mat = (Matrix)_form.Items.Item(FormIds.MatBill).Specific;
                mat.Columns.Item("colQty").Cells.Item(1).Click(BoCellClickType.ct_Regular, 0);
            }
            catch
            {
                // confort uniquement
            }
        }

        private void BuildMatrix()
        {
            DataTable dt = _form.DataSources.DataTables.Add(Dt);
            dt.Columns.Add("Lbl", BoFieldsType.ft_AlphaNumeric, 50);
            dt.Columns.Add("Val", BoFieldsType.ft_Sum);
            dt.Columns.Add("Qty", BoFieldsType.ft_Integer);
            dt.Columns.Add("Amt", BoFieldsType.ft_Sum);

            dt.Rows.Add(_denoms.Count);
            for (int i = 0; i < _denoms.Count; i++)
            {
                dt.SetValue("Lbl", i, _denoms[i].Name);
                dt.SetValue("Val", i, _denoms[i].Value);
                dt.SetValue("Qty", i, 0);
                dt.SetValue("Amt", i, 0.0);
            }

            Item item = _form.Items.Add(FormIds.MatBill, BoFormItemTypes.it_MATRIX);
            item.Left = 10; item.Top = 62; item.Width = 420; item.Height = 300;
            Matrix mat = (Matrix)item.Specific;
            mat.SelectionMode = BoMatrixSelect.ms_None;

            AddColumn(mat, "colLbl", "Coupure", 140, "Lbl", false);
            AddColumn(mat, "colVal", "Valeur", 80, "Val", false);
            AddColumn(mat, "colQty", "Quantité", 70, "Qty", true);
            AddColumn(mat, "colAmt", "Montant", 100, "Amt", false);
            mat.LoadFromDataSource();
        }

        private static void AddColumn(Matrix mat, string id, string caption, int width, string boundColumn, bool editable)
        {
            Column col = mat.Columns.Add(id, BoFormItemTypes.it_EDIT);
            col.TitleObject.Caption = caption;
            col.Width = width;
            col.Editable = editable;
            col.DataBind.Bind(Dt, boundColumn);
        }

        private void App_ItemEvent(string formUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            if (formUID != FormIds.BillForm)
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
                if (pVal.EventType == BoEventTypes.et_VALIDATE && pVal.ItemUID == FormIds.MatBill &&
                    pVal.ColUID == "colQty" && pVal.ItemChanged && pVal.Row > 0)
                {
                    RecalcLine(pVal.Row);
                    return;
                }

                if (pVal.EventType != BoEventTypes.et_ITEM_PRESSED || !pVal.ActionSuccess)
                    return;

                if (pVal.ItemUID == FormIds.BtnBillCancel)
                    _form.Close();
                else if (pVal.ItemUID == FormIds.BtnBillOk)
                    HandleOk();
            }
            catch (Exception ex)
            {
                _app.MessageBox(ex.Message);
            }
        }

        /// <summary>Recalcule le montant d'une ligne après saisie de sa quantité, puis les totaux.</summary>
        private void RecalcLine(int row)
        {
            Matrix mat = (Matrix)_form.Items.Item(FormIds.MatBill).Specific;
            DataTable dt = _form.DataSources.DataTables.Item(Dt);
            mat.GetLineData(row);

            int qty = Convert.ToInt32(dt.GetValue("Qty", row - 1));
            if (qty < 0)
            {
                qty = 0;
                dt.SetValue("Qty", row - 1, 0);
                _app.StatusBar.SetText("La quantité ne peut pas être négative.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
            }
            dt.SetValue("Amt", row - 1, qty * _denoms[row - 1].Value);
            mat.SetLineData(row);
            UpdateTotals();
        }

        private List<CountLine> ReadLines()
        {
            DataTable dt = _form.DataSources.DataTables.Item(Dt);
            var lines = new List<CountLine>();
            for (int i = 0; i < _denoms.Count; i++)
            {
                lines.Add(new CountLine
                {
                    DenomCode = _denoms[i].Code,
                    Label = _denoms[i].Name,
                    Value = _denoms[i].Value,
                    Qty = Math.Max(0, Convert.ToInt32(dt.GetValue("Qty", i)))
                });
            }
            return lines;
        }

        private void UpdateTotals()
        {
            double total = ReadLines().Sum(l => l.Amount);
            double diff = total - _expected;
            UserDataSources uds = _form.DataSources.UserDataSources;
            uds.Item(UdTotal).ValueEx = CashSessionService.FormatAmount(total);
            uds.Item(UdDiff).ValueEx = CashSessionService.FormatAmount(diff) + (diff > 0 ? " (excédent)" : diff < 0 ? " (manquant)" : "");
        }

        private void HandleOk()
        {
            Matrix mat = (Matrix)_form.Items.Item(FormIds.MatBill).Specific;
            mat.FlushToDataSource();
            List<CountLine> lines = ReadLines();
            if (lines.Any(l => l.Qty < 0))
            {
                _app.MessageBox("Une quantité ne peut pas être négative.");
                return;
            }

            double total = lines.Sum(l => l.Amount);
            double diff = total - _expected;
            int answer = _app.MessageBox(
                "Montant attendu : " + CashSessionService.FormatAmount(_expected) + "\n" +
                "Total compté : " + CashSessionService.FormatAmount(total) + "\n" +
                "Écart : " + CashSessionService.FormatAmount(diff) + (diff > 0 ? " (excédent)" : diff < 0 ? " (manquant)" : "") +
                (Math.Abs(diff) > 0.0001 ? "\nL'écart sera comptabilisé sur le compte d'écarts de caisse." : "") + "\n\n" +
                _confirmQuestion, 2, "Oui", "Non");
            if (answer != 1)
                return;

            // Une exception remonte au gestionnaire d'événement : la fenêtre reste ouverte
            _onConfirmed(lines);

            _form.Close();
            try
            {
                _onSaved?.Invoke();
            }
            catch (Exception ex)
            {
                _app.MessageBox("Le billetage est bien enregistré, mais l'écran n'a pas pu être actualisé : " + ex.Message);
            }
        }

        // ---------- Helpers ----------

        private void AddLabel(string id, string caption, int left, int top, string linkTo)
        {
            Item item = _form.Items.Add(id, BoFormItemTypes.it_STATIC);
            item.Left = left; item.Top = top; item.Width = 140; item.Height = 15;
            item.LinkTo = linkTo;
            ((StaticText)item.Specific).Caption = caption;
        }

        private void AddAmountField(string id, int left, int top, string udsId)
        {
            Item item = _form.Items.Add(id, BoFormItemTypes.it_EDIT);
            item.Left = left; item.Top = top; item.Width = 160; item.Height = 19;
            item.RightJustified = true;
            ((EditText)item.Specific).DataBind.SetBound(true, "", udsId);
            item.Enabled = false;
        }

        private void AddButton(string id, string caption, int left, int top, int width)
        {
            Item item = _form.Items.Add(id, BoFormItemTypes.it_BUTTON);
            item.Left = left; item.Top = top; item.Width = width; item.Height = 22;
            ((Button)item.Specific).Caption = caption;
        }
    }
}
