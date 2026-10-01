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

        public static void CreateMenus(Application app)
        {
            if (app.Menus.Exists(FormIds.MenuRoot))
                return;

            MenuCreationParams p = (MenuCreationParams)app.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);
            p.Type = BoMenuType.mt_POPUP;
            p.UniqueID = FormIds.MenuRoot;
            p.String = "Petty Cash";
            p.Position = -1;
            MenuItem root = app.Menus.AddEx(p);

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
    }
}
