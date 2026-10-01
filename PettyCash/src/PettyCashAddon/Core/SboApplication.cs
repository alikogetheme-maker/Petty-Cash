using SAPbouiCOM;
using PettyCashAddon.Forms;

namespace PettyCashAddon.Core
{
    /// <summary>
    /// Connexion à l'application UI API et création du menu "Petty Cash".
    /// </summary>
    internal static class SboApplication
    {
        public static Application Connect(string connectionString)
        {
            SboGuiApi guiApi = new SboGuiApi();
            guiApi.Connect(connectionString);
            return guiApi.GetApplication(-1);
        }

        // Menu standard "Modules" de SAP B1 (menu principal à gauche)
        private const string ModulesMenuUid = "43520";

        public static void CreateMenus(Application app)
        {
            if (app.Menus.Exists(FormIds.MenuRoot))
                app.Menus.RemoveEx(FormIds.MenuRoot);

            MenuCreationParams p = (MenuCreationParams)app.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);
            p.Type = BoMenuType.mt_POPUP;
            p.UniqueID = FormIds.MenuRoot;
            p.String = "Petty Cash";
            p.Position = -1;
            string image = ExtractMenuImage();
            if (image != null)
                p.Image = image;

            MenuItem root;
            try
            {
                root = app.Menus.Item(ModulesMenuUid).SubMenus.AddEx(p);
            }
            catch (System.Exception ex) when (image != null)
            {
                // Image refusée par le client : on crée le menu sans icône
                Program.Log("Menu avec icône refusé, création sans icône : " + ex.Message);
                p.Image = "";
                root = app.Menus.Item(ModulesMenuUid).SubMenus.AddEx(p);
            }

            p = (MenuCreationParams)app.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);
            p.Type = BoMenuType.mt_STRING;
            p.UniqueID = FormIds.MenuSession;
            p.String = "Session de caisse";
            root.SubMenus.AddEx(p);

            p = (MenuCreationParams)app.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);
            p.Type = BoMenuType.mt_STRING;
            p.UniqueID = FormIds.MenuReport;
            p.String = "Rapport de caisse";
            root.SubMenus.AddEx(p);
        }

        /// <summary>
        /// SAP attend un chemin de fichier BMP : on extrait l'icône intégrée
        /// à l'exe dans le dossier temporaire. Sans icône, le menu est créé quand même.
        /// </summary>
        private static string ExtractMenuImage()
        {
            try
            {
                string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "PettyCashMenu.bmp");
                using (var res = typeof(SboApplication).Assembly.GetManifestResourceStream("PettyCashAddon.Resources.PettyCash.bmp"))
                {
                    if (res == null)
                        return null;
                    using (var file = System.IO.File.Create(path))
                        res.CopyTo(file);
                }
                return path;
            }
            catch (System.Exception ex)
            {
                Program.Log("Icône du menu non extraite : " + ex.Message);
                return null;
            }
        }
    }
}
