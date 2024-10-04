using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using AddOnUI.Logica;
using AddOnUI.DB_Structure;
using System.Reflection;

namespace AddOnUI.App
{
    public class Main
    {
        public Main()
        {
            Connect.SetApplication();
            Connect.ConnectToCompany();

            Connect.SetFilters();
            Globals.SAPVersion = Globals.oCompany.Version;
            Globals.SBO_Application.SetStatusBarMessage("Validando estructura de la Base de Datos", SAPbouiCOM.BoMessageTime.bmt_Short, false);

            Globals.Addon = Assembly.GetEntryAssembly().GetName().Name;
            Version version = Assembly.GetEntryAssembly().GetName().Version;
            Globals.version = version.ToString().Replace(".0.0", "");

            #region Estructura
            Setup oSetup = new Setup();
            Globals.Actual = oSetup.validarVersion(Globals.Addon, Globals.version);
            if (Globals.Actual == false)
            {
                CreateStructure.CreateStruct();
                oSetup.confirmarVersion(Globals.Addon, Globals.version);
                oSetup.confirmarVersionUpdate(Globals.Addon, Globals.version);
                Globals.continuar = 0;
            }
            else
                Globals.continuar = 0;
            #endregion

            Globals.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
            Globals.SBO_Application.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormDataEvent);
            Globals.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
            Globals.SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);

            Menu.AddMenu();
            Menu.AddMenuItems();

            Globals.continuar = 0;
            Globals.SBO_Application.MessageBox("Addon Migrar OfV conectado");
        }


        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.FormTypeEx != "0")
            {
                try
                {
                    SAPbouiCOM.Form oForm = Globals.SBO_Application.Forms.Item(pVal.FormUID);

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                    {
                        if (pVal.Before_Action)
                        {
                            if (pVal.FormType == 149)
                            {
                                OfertaVenta.GenerarControles(oForm);
                            }
                            //if (pVal.FormType == 1470000200)
                            //{
                            //    FacturaVenta.GenerarControles(oForm);
                            //}                          
                        }
                        else
                        {
                            //despues de 
                        }
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                    {
                        if (pVal.Before_Action)
                        {
                            if ((pVal.FormType == 133 || pVal.FormType == 141) && pVal.ItemUID == "1" && (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE))
                            {
                                Validaciones.Validar(oForm,pVal,pVal.FormType);
                            }

                        }
                        else
                        {
                            if (pVal.FormType == 149 && pVal.ItemUID == "BtnMig" && oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                
                                OfertaVenta.LineasOfertaVenta(oForm, pVal);
                            }

                            //if (pVal.FormType == 133 && pVal.ItemUID == "BtnAnu" && oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            //{
                            //    FacturaVenta.AnularVenta(oForm);
                            //}

                            //if (pVal.FormType == 141 && pVal.ItemUID == "BtnAnu" && oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            //{
                            //    FacturaCompra.AnularCompra(oForm);
                            //}

                            //if (pVal.FormType == 721 && pVal.ItemUID == "BtnAnu" && oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            //{
                            //    EntradaMercancia.AnularEntrada(oForm);
                            //}

                            //if (pVal.FormTypeEx == "PruebasForm" && pVal.ItemUID == "Item_8")
                            //{
                            //    FormularioPrueba.Cerrar(oForm);
                            //}
                            //if (pVal.FormTypeEx == "PruebasForm" && pVal.ItemUID == "Item_4")
                            //{
                            //    FormularioPrueba.Cargar(oForm);
                            //}
                            //if (pVal.FormTypeEx == "PruebasForm" && pVal.ItemUID == "Item_7")
                            //{
                            //    FormularioPrueba.AnularFV(oForm);
                            //}
                        }
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED)
                    {
                        if (pVal.Before_Action)
                        {

                        }
                        else
                        {
                            if (pVal.FormTypeEx == "PruebasForm" && pVal.ItemUID == "Item_10" )
                            {
                                FormularioPrueba.LinkPressed(oForm, pVal);
                            }
                        }

                    }

                }
                catch (Exception ex)
                {
                    BubbleEvent = false;
                    if (ex.Message.IndexOf("Form - Not found  [66000-9]") != -1)
                    {
                        Globals.Error = "TP: Activar campos de usuario al crear un documento";
                        Globals.SBO_Application.SetStatusBarMessage(Globals.Error, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    }
                    else
                    {
                        Globals.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    }
                }
            }
        }

        private void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.Form oForm = Globals.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID);

                if ((BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) & BusinessObjectInfo.ActionSuccess == true)
                {
                    if (BusinessObjectInfo.FormTypeEx == "149")
                    {
                        string ObjectKey = BusinessObjectInfo.ObjectKey;
                        OfertaVenta.ActualizarOferta(ObjectKey);
                    }
                }

            }
            catch (Exception ex)
            {
                Globals.SBO_Application.SetStatusBarMessage(ex.Message);
                BubbleEvent = false;
                return;
            }
        }

        private void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            if (EventType == SAPbouiCOM.BoAppEventTypes.aet_ShutDown)
            {
                System.Windows.Forms.Application.Exit();
            }
            if (EventType == SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged)
            {
                System.Windows.Forms.Application.Exit();
            }
            if (EventType == SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged)
            {
                System.Windows.Forms.Application.Exit();
            }
            if (EventType == SAPbouiCOM.BoAppEventTypes.aet_FontChanged)
            {
                System.Windows.Forms.Application.Exit();
            }
            if (EventType == SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition)
            {
                System.Windows.Forms.Application.Exit();
            }
        }

        private void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.BeforeAction)
            {
            }
            else
            {
                switch (pVal.MenuUID)
                {
                    case "SM_ADDON_PRUEBAS":
                        FormularioPrueba.CargarFormulario();
                        break;
                }
            }
        }
    }
}
