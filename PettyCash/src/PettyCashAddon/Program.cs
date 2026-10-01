using System;
using SAPbouiCOM;
using PettyCashAddon.Core;
using PettyCashAddon.Setup;
using PettyCashAddon.Forms;

namespace PettyCashAddon
{
    internal static class Program
    {
        // NB : ne pas faire "using System.Windows.Forms;" ici, ce namespace
        // définit aussi un type "Application" qui entrerait en conflit avec
        // SAPbouiCOM.Application. On qualifie donc explicitement ci-dessous.

        private static Application _sboApplication;
        private static CashSessionFormController _sessionController;
        private static CashReportFormController _reportController;

        [STAThread]
        private static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Cet exécutable doit être lancé par SAP Business One (Add-On Administration), pas directement.",
                    "PettyCashAddon");
                return;
            }

            try
            {
                _sboApplication = SboApplication.Connect(args[0]);
                DiCompany.Connect(_sboApplication);
                MetadataSetup.EnsureSchema();
                SboApplication.CreateMenus(_sboApplication);

                _sessionController = new CashSessionFormController(_sboApplication);
                _reportController = new CashReportFormController(_sboApplication);

                _sboApplication.MenuEvent += SboApplication_MenuEvent;
                _sboApplication.AppEvent += SboApplication_AppEvent;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Erreur au démarrage de l'add-on Petty Cash : " + ex.Message, "PettyCashAddon");
                return;
            }

            System.Windows.Forms.Application.Run();
        }

        private static void SboApplication_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.BeforeAction)
                return;

            try
            {
                if (pVal.MenuUID == FormIds.MenuSession)
                    _sessionController.ShowOrActivate();
                else if (pVal.MenuUID == FormIds.MenuReport)
                    _reportController.ShowOrActivate();
            }
            catch (Exception ex)
            {
                _sboApplication.MessageBox(ex.Message);
            }
        }

        private static void SboApplication_AppEvent(BoAppEventTypes EventType)
        {
            if (EventType == BoAppEventTypes.aet_ShutDown || EventType == BoAppEventTypes.aet_CompanyChanged)
                System.Windows.Forms.Application.Exit();
        }
    }
}
