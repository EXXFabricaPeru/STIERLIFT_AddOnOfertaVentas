using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AddOnUI.App
{
    class Connect
    {
        public static void SetApplication()
        {
            try
            {
                SAPbouiCOM.SboGuiApi SboGuiApi = default(SAPbouiCOM.SboGuiApi);
                string sConnectionString = null;
                SboGuiApi = new SAPbouiCOM.SboGuiApi();
                if (Environment.GetCommandLineArgs().Length > 1)
                {
                    sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
                }
                else
                {
                    sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(0));
                }
                SboGuiApi.Connect(sConnectionString);
                Globals.SBO_Application = SboGuiApi.GetApplication();
            }
            catch (Exception ex)
            {
                Globals.SBO_Application.MessageBox(ex.Message);
            }
        }

        public static bool ConnectToCompany()
        {
            try
            {
                Globals.oCompany = Globals.SBO_Application.Company.GetDICompany();
                return true;
            }
            catch (Exception ex)
            {
                Globals.SBO_Application.MessageBox(ex.Message);
                return false;
            }
        }

        public static void SetFilters()
        {
            Globals.oFilters = new SAPbouiCOM.EventFilters();
            
            Globals.oFilter = Globals.oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK);
            Globals.SBO_Application.SetFilter(Globals.oFilters);

            #region COMBO_SELECT
            Globals.oFilter = Globals.oFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT);
            
            Globals.oFilter.AddEx("PruebasForm");
            Globals.SBO_Application.SetFilter(Globals.oFilters);
            #endregion

            #region ITEM_PRESSED
            Globals.oFilter = Globals.oFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED);
            //Globals.oFilter.AddEx("PruebasForm");
            Globals.oFilter.AddEx("149");//Oferta de Ventas
            //Globals.oFilter.AddEx("141");//Factura de Compras
            //Globals.oFilter.AddEx("721");//Entrada de inventario
            //Globals.oFilter.AddEx("426");//Pagos Efectuados
            Globals.SBO_Application.SetFilter(Globals.oFilters);
            #endregion

            #region MATRIX_LINK_PRESSED
            Globals.oFilter = Globals.oFilters.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED);
            Globals.oFilter.AddEx("PruebasForm");
            Globals.SBO_Application.SetFilter(Globals.oFilters);
            #endregion

            #region CLICK
            Globals.oFilter = Globals.oFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK);
            Globals.oFilter.AddEx("149");//Oferta de Ventas
            //Globals.oFilter.AddEx("133");//Factura de Ventas
            //Globals.oFilter.AddEx("139");//OV
            //Globals.oFilter.AddEx("142");//Pedido
            Globals.oFilter.AddEx("1470000200"); //Requerimiento
            Globals.SBO_Application.SetFilter(Globals.oFilters);
            #endregion

            #region FORM_LOAD
            Globals.oFilter = Globals.oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD);
            Globals.oFilter.AddEx("149");//Oferta de Ventas
            //Globals.oFilter.AddEx("133");//Factura de Ventas
            //Globals.oFilter.AddEx("179");//Nota de Credito de Ventas
            //Globals.oFilter.AddEx("139");//OV
            //Globals.oFilter.AddEx("141");//Factura de Compras
            //Globals.oFilter.AddEx("142");//Pedido
            //Globals.oFilter.AddEx("721");//Entrada de inventario
            //Globals.oFilter.AddEx("426");//Pagos Efectuados
            Globals.SBO_Application.SetFilter(Globals.oFilters);
            #endregion

            #region FORM_DATA_LOAD
            Globals.oFilter = Globals.oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD);
            Globals.oFilter.AddEx("149");//Oferta de Ventas
            //Globals.oFilter.AddEx("133");//Factura de Ventas
            //Globals.oFilter.AddEx("139");//OV
            //Globals.oFilter.AddEx("142");//Pedido
            Globals.oFilter.AddEx("1470000200"); //Requerimiento
            Globals.SBO_Application.SetFilter(Globals.oFilters);
            #endregion

            #region FORM_DATA_ADD
            Globals.oFilter = Globals.oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD);
            Globals.oFilter.AddEx("1470000200"); //Requerimiento
            Globals.oFilter.AddEx("149");//Oferta de Ventas
            //Globals.oFilter.AddEx("133"); //Factura de Ventas
            Globals.SBO_Application.SetFilter(Globals.oFilters);
            #endregion

            #region FORM_DATA_UPDATE
            Globals.oFilter = Globals.oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE);
            Globals.oFilter.AddEx("149");//Oferta de Ventas
            //Globals.oFilter.AddEx("134"); //Socio de Negocio
            Globals.SBO_Application.SetFilter(Globals.oFilters);
            #endregion

        }
    }
}
