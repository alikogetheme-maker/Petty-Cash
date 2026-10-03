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

        /// <summary>Journal de démarrage : %TEMP%\PettyCashAddon.log</summary>
        internal static readonly string LogPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "PettyCashAddon.log");

        /// <summary>
        /// Chaîne de connexion de développement fournie par SAP : permet de
        /// lancer l'add-on depuis Visual Studio (F5) sur le client SAP B1 déjà
        /// ouvert, sans passer par la gestion des add-ons.
        /// </summary>
        private const string DevConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";

        [STAThread]
        private static void Main(string[] args)
        {
            if (args.Length == 0 && System.Diagnostics.Debugger.IsAttached)
                args = new[] { DevConnectionString };

            // Main ne référence aucun type SAP : une DLL manquante est ainsi
            // interceptée ici (et journalisée) au lieu de faire planter le
            // processus avant même le premier try.
            AppDomain.CurrentDomain.UnhandledException += (s, e) => Log("Exception non gérée : " + e.ExceptionObject);
            Log("Démarrage " + typeof(Program).Assembly.Location + " (" + (Environment.Is64BitProcess ? "64" : "32") + " bits), " + args.Length + " argument(s)");

            if (args.Length == 0)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Cet exécutable doit être lancé par SAP Business One (Add-On Administration), pas directement.",
                    "PettyCashAddon");
                return;
            }

            try
            {
                Start(args[0]);
            }
            catch (Exception ex)
            {
                Log("Échec du démarrage : " + ex);
                System.Windows.Forms.MessageBox.Show("Erreur au démarrage de l'add-on Petty Cash : " + ex.Message + "\n\nDétail : " + LogPath, "PettyCashAddon");
                return;
            }

            Log("Add-on prêt");
            System.Windows.Forms.Application.Run();
        }

        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
        private static void Start(string connectionString)
        {
            _sboApplication = SboApplication.Connect(connectionString);
            Log("UI API connectée");
            DiCompany.Connect(_sboApplication);
            Log("DI API connectée");
            MetadataSetup.EnsureSchema();
            Log("Tables et champs vérifiés");
            SboApplication.CreateMenus(_sboApplication);

            // Écran de billetage partagé par la session (ouverture / clôture) et le rapport (clôture)
            var billForm = new BillCountForm(_sboApplication);
            _sessionController = new CashSessionFormController(_sboApplication, billForm);
            _reportController = new CashReportFormController(_sboApplication, billForm, _sessionController.RefreshIfOpen);

            _sboApplication.MenuEvent += SboApplication_MenuEvent;
            _sboApplication.AppEvent += SboApplication_AppEvent;
            _sboApplication.StatusBar.SetText("Add-on Petty Cash démarré.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
        }

        internal static void Log(string message)
        {
            try
            {
                System.IO.File.AppendAllText(LogPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + message + Environment.NewLine);
            }
            catch
            {
                // le journal ne doit jamais empêcher l'add-on de fonctionner
            }
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
            if (EventType == BoAppEventTypes.aet_ShutDown ||
                EventType == BoAppEventTypes.aet_CompanyChanged ||
                EventType == BoAppEventTypes.aet_ServerTerminition)
            {
                try
                {
                    if (_sboApplication.Menus.Exists(FormIds.MenuRoot))
                        _sboApplication.Menus.RemoveEx(FormIds.MenuRoot);
                }
                catch
                {
                    // le client est peut-être déjà fermé
                }
                DiCompany.Disconnect();
                System.Windows.Forms.Application.Exit();
            }
        }
    }
}
