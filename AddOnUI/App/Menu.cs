using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AddOnUI.App
{
    public class Menu
    {
        public static void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenus = default(SAPbouiCOM.Menus);
            SAPbouiCOM.MenuItem oMenuItem = default(SAPbouiCOM.MenuItem);
            oMenus = Globals.SBO_Application.Menus;
            SAPbouiCOM.MenuCreationParams oCreationPackage = default(SAPbouiCOM.MenuCreationParams);
            oCreationPackage = 
                Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);

            try
            {
                oMenuItem = Globals.SBO_Application.Menus.Item("2048");
                oMenus = oMenuItem.SubMenus;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "SM_ADDON_PRUEBAS";
                oCreationPackage.String = "Archivo AddOn Pruebas";
                oCreationPackage.Position = oMenuItem.SubMenus.Count + 1;

                if (!(oMenus.Exists("SM_ADDON_PRUEBAS")))
                {
                    oMenus.AddEx(oCreationPackage);
                }

                oMenuItem = Globals.SBO_Application.Menus.Item("43537");
                oMenus = oMenuItem.SubMenus;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "SM_ADDON_BANCO";
                oCreationPackage.String = "Pagos Proveedores Masivo";
                oCreationPackage.Position = oMenuItem.SubMenus.Count + 1;

                if (!(oMenus.Exists("SM_ADDON_BANCO")))
                {
                    oMenus.AddEx(oCreationPackage);
                }
            }
            catch (Exception ex)
            {
                Globals.SBO_Application.SetStatusBarMessage(ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, false);
            }
        }

        public static void AddMenu()
        {
            SAPbouiCOM.Menus oMenus = default(SAPbouiCOM.Menus);
            SAPbouiCOM.MenuItem oMenuItem = default(SAPbouiCOM.MenuItem);
            oMenus = Globals.SBO_Application.Menus;
            SAPbouiCOM.MenuCreationParams oCreationPackage = default(SAPbouiCOM.MenuCreationParams);
            oCreationPackage =
                Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);

            try
            {
                oMenuItem = Globals.SBO_Application.Menus.Item("43520");
                oMenus = oMenuItem.SubMenus;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                oCreationPackage.UniqueID = "SM_SBA";
                oCreationPackage.String = "Modulo SBA";
                oCreationPackage.Position = oMenuItem.SubMenus.Count + 1;

                //oCreationPackage.Image = ""//ruta iamgen

                if (!(oMenus.Exists("SM_SBA")))
                {
                    oMenus.AddEx(oCreationPackage);
                }

                oMenuItem = Globals.SBO_Application.Menus.Item("SM_SBA");
                oMenus = oMenuItem.SubMenus;

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                oCreationPackage.UniqueID = "SM_SBA_FUN";
                oCreationPackage.String = "SBA Funcionalidades";
                oMenus.AddEx(oCreationPackage);

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                oCreationPackage.UniqueID = "SM_SBA_REP";
                oCreationPackage.String = "SBA Reportes";
                oMenus.AddEx(oCreationPackage);

                //level 2

                oMenuItem = Globals.SBO_Application.Menus.Item("SM_SBA_FUN");
                oMenus = oMenuItem.SubMenus;

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "SM_SBA_FUN1";
                oCreationPackage.String = "SBA Funcionalidades 1";
                oMenus.AddEx(oCreationPackage);

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "SM_SBA_FUN2";
                oCreationPackage.String = "SBA Funcionalidades 2";
                oMenus.AddEx(oCreationPackage);


                //level 2

                oMenuItem = Globals.SBO_Application.Menus.Item("SM_SBA_REP");
                oMenus = oMenuItem.SubMenus;

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "SM_SBA_REP1";
                oCreationPackage.String = "SBA Reporte 1";
                oMenus.AddEx(oCreationPackage);

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "SM_SBA_REP2";
                oCreationPackage.String = "SBA Reporte 2";
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception ex)
            {
                Globals.SBO_Application.SetStatusBarMessage(ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, false);
            }
        }
    }
}

