using AddOnUI.App;
using AddOnUI.Modelos;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AddOnUI.Logica
{
    public class OfertaVenta
    {
        public static void GenerarControles(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.Button obtnCancel = oForm.Items.Item("2").Specific;


                SAPbouiCOM.Item oItem = oForm.Items.Add("BtnMig", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = obtnCancel.Item.Left + obtnCancel.Item.Width + 7;//600;
                oItem.Top = obtnCancel.Item.Top;//125;
                oItem.Width = 100;
                oItem.Height = obtnCancel.Item.Height;

                SAPbouiCOM.Button oBtn = ((SAPbouiCOM.Button)(oItem.Specific));
                oBtn.Caption = "Migrar a SC";
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void ActualizarOferta(string ObjectKey)
        {
            try
            {
                string KeyXML = ObjectKey.Substring(ObjectKey.IndexOf("<DocEntry>") + 10, ObjectKey.IndexOf("</DocEntry>") - (ObjectKey.IndexOf("<DocEntry>") + 10));

                Globals.Query = "UPDATE QUT1 SET U_EXX_OFBA = \"DocEntry\",U_EXX_LOFB = \"LineNum\" WHERE \"DocEntry\" = " + KeyXML;
                Globals.RunQuery(Globals.Query);
                Globals.Release(Globals.oRec);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }


        #region anterior
        /*
        public static void LineasOfertaVenta(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal)
        {
            try 
            {
                //SAPbouiCOM.Form oUserForm = Globals.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                int iValor = Globals.SBO_Application.MessageBox("Se procederá a copiar la información selecionada a una solicitud de compra.\nDesea continuar?", 2, "Si", "No");

                if (iValor == 2)
                {
                    return;
                }


                SAPbouiCOM.EditText oDocNum = oForm.Items.Item("8").Specific;


                if (oDocNum.Value != null || oDocNum.Value != "")
                {
                    Globals.Query = "SELECT \"DocEntry\" FROM OQUT WHERE \"DocNum\"  ='" + oDocNum.Value.ToString() + "' ";
                    Globals.RunQuery(Globals.Query);
                    int iDocEntry = Int32.Parse(Globals.oRec.Fields.Item(0).Value.ToString());
                    Globals.Release(Globals.oRec);


                    SAPbobsCOM.Documents oSalesQuotation;
                    oSalesQuotation = (SAPbobsCOM.Documents)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQuotations);
                    if (oSalesQuotation.GetByKey(iDocEntry))
                    {
                        string Aprob = oSalesQuotation.UserFields.Fields.Item("U_EXS_APROB").Value;
                        if (Aprob == "-")
                        {

                            OfertaVentaDatos Doc = new OfertaVentaDatos();
                            Doc.DocNum = Convert.ToInt32(oDocNum.Value);
                            Doc.DocEntry = iDocEntry;
                            Doc.CardCode = oSalesQuotation.CardCode;
                            Doc.DocDate = oSalesQuotation.DocDate;
                            Doc.DocDueDate = oSalesQuotation.DocDueDate;
                            Doc.OT = oSalesQuotation.UserFields.Fields.Item("U_vOt").Value;
                            Doc.Taller = oSalesQuotation.UserFields.Fields.Item("U_vTallCode").Value;
                            Doc.Sucursal = oSalesQuotation.BPL_IDAssignedToInvoice.ToString();
                            SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("38").Specific; // matrix de detalles

                            if (oMatrix.RowCount > 0)
                            {
                                for (int i = 1; i <= oMatrix.RowCount; i++)
                                {
                                    if (oMatrix.IsRowSelected(i))
                                    {
                                        oSalesQuotation.Lines.SetCurrentLine(i-1);
                                        string ItemCode = oSalesQuotation.Lines.ItemCode;
                                        double cant = oSalesQuotation.Lines.Quantity;
                                        //string newString = String.Format(AddOnUI.Properties.Resources.ItemValidate, ItemCode, iDocEntry);
                                        Globals.Query = String.Format(AddOnMigrarOfV.Properties.Resources.ItemValidate, iDocEntry, i);
                                        Globals.RunQuery(Globals.Query);
                                        double sum = double.Parse(Globals.oRec.Fields.Item(0).Value.ToString());
                                        Globals.Release(Globals.oRec);
                                        if (sum < cant)
                                        {
                                            DetalleOV detalle = new DetalleOV();
                                            detalle.ItemCode = ItemCode;
                                            detalle.Descripcion = oSalesQuotation.Lines.ItemDescription;
                                            detalle.Cantidad = cant-sum;
                                            detalle.Precio = oSalesQuotation.Lines.UnitPrice;
                                            detalle.TaxCode = oSalesQuotation.Lines.TaxCode;
                                            detalle.UoM = oSalesQuotation.Lines.UoMCode;
                                            detalle.Almacen = oSalesQuotation.Lines.WarehouseCode;
                                            detalle.Margen = oSalesQuotation.Lines.CostingCode;
                                            detalle.CentroCosto = oSalesQuotation.Lines.CostingCode2;
                                            detalle.CentroBeneficio = oSalesQuotation.Lines.CostingCode3;
                                            detalle.Equipos = oSalesQuotation.Lines.CostingCode4;
                                            detalle.CC5 = oSalesQuotation.Lines.CostingCode5;
                                            detalle.Proyecto = oSalesQuotation.Lines.ProjectCode;
                                            detalle.GrupoDetraccion = oSalesQuotation.Lines.UserFields.Fields.Item("U_EXX_GRUPODET").Value;
                                            detalle.GrupoRetencion = oSalesQuotation.Lines.UserFields.Fields.Item("U_EXX_GRUPOPER").Value;
                                            detalle.NroOferta = iDocEntry.ToString();
                                            detalle.LineaOferta = i.ToString();

                                            Doc.Lineas.Add(detalle);
                                        }
                                        else
                                        {
                                            Globals.SBO_Application.SetStatusBarMessage("La linea seleccionada con el Item " + ItemCode + " ya fue asignado a otra Solicitud de Compra", SAPbouiCOM.BoMessageTime.bmt_Short, true);

                                        }
                                    }
                                }
                            }
                            PurchaseRequest(Doc);
                        }                       
                    }
                }
                else
                {
                    Globals.SBO_Application.SetStatusBarMessage("Debe seleccionar un Documento", SAPbouiCOM.BoMessageTime.bmt_Short, false);

                }

            }
            catch (Exception ex)
            {
                throw ex;
            }


        }
        */
#endregion



        public static void LineasOfertaVenta(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                //SAPbouiCOM.Form oUserForm = Globals.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                int iValor = Globals.SBO_Application.MessageBox("Se procederá a copiar la información selecionada a una solicitud de compra.\nDesea continuar?", 2, "Si", "No");

                if (iValor == 2)
                {
                    return;
                }


                SAPbouiCOM.EditText oDocNum = oForm.Items.Item("8").Specific;


                if (oDocNum.Value != null || oDocNum.Value != "")
                {
                    Globals.Query = "SELECT \"DocEntry\" FROM OQUT WHERE \"DocNum\"  ='" + oDocNum.Value.ToString() + "' ";
                    Globals.RunQuery(Globals.Query);
                    int iDocEntry = Int32.Parse(Globals.oRec.Fields.Item(0).Value.ToString());
                    Globals.Release(Globals.oRec);


                    SAPbobsCOM.Documents oSalesQuotation;
                    oSalesQuotation = (SAPbobsCOM.Documents)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQuotations);
                    bool band = oSalesQuotation.GetByKey(iDocEntry);
                    if (band)
                    {
                        string Aprob = oSalesQuotation.UserFields.Fields.Item("U_EXS_APROB").Value;
                        if (Aprob == "A")
                        {

                            OfertaVentaDatos Doc = new OfertaVentaDatos();
                            Doc.DocNum = Convert.ToInt32(oDocNum.Value);
                            Doc.DocEntry = iDocEntry;
                            Doc.CardCode = oSalesQuotation.CardCode;
                            Doc.DocDate = oSalesQuotation.DocDate;
                            Doc.DocDueDate = oSalesQuotation.DocDueDate;
                            Doc.OT = oSalesQuotation.UserFields.Fields.Item("U_vOt").Value;
                            Doc.Taller = oSalesQuotation.UserFields.Fields.Item("U_vTallCode").Value;
                            Doc.Sucursal = oSalesQuotation.BPL_IDAssignedToInvoice.ToString();
                            SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("38").Specific; // matrix de detalles

                            if (oMatrix.RowCount > 0)
                            {
                                for (int i = 1; i <= oMatrix.RowCount; i++)
                                {
                                    if (oMatrix.IsRowSelected(i))
                                    {
                                        //oSalesQuotation.Lines.SetCurrentLine(i - 1);
                                        string ItemCode = oMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.ToString();//oSalesQuotation.Lines.ItemCode;
                                        if (ItemCode == "")
                                        {
                                            Globals.SBO_Application.SetStatusBarMessage("Línea " + i.ToString() + " no disponible para selección", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                        }
                                        else
                                        {
                                            double cant = double.Parse(oMatrix.Columns.Item("11").Cells.Item(i).Specific.Value.ToString());// oSalesQuotation.Lines.Quantity;
                                            int iLinea = Int32.Parse(oMatrix.Columns.Item("U_EXX_LOFB").Cells.Item(i).Specific.Value.ToString());

                                            Globals.Query = String.Format(AddOnMigrarOfV.Properties.Resources.ItemValidate, iDocEntry, iLinea);
                                            Globals.RunQuery(Globals.Query);
                                            double sum = double.Parse(Globals.oRec.Fields.Item(0).Value.ToString());
                                            Globals.Release(Globals.oRec);
                                            if (sum < cant)
                                            {
                                                DetalleOV detalle = new DetalleOV();
                                                detalle.ItemCode = ItemCode;
                                                detalle.Descripcion = oMatrix.Columns.Item("3").Cells.Item(i).Specific.Value.ToString();// oSalesQuotation.Lines.ItemDescription;
                                                detalle.Cantidad = cant - sum;
                                                string sPrecio = oMatrix.Columns.Item("U_EXS_CU").Cells.Item(i).Specific.Value.ToString();
                                                string[] ArrayPrecio = sPrecio.Split(' ');
                                                detalle.Precio = double.Parse(ArrayPrecio[0].ToString());// oSalesQuotation.Lines.UnitPrice;
                                                detalle.TaxCode = oMatrix.Columns.Item("160").Cells.Item(i).Specific.Value.ToString();// oSalesQuotation.Lines.TaxCode;
                                                detalle.UoM = oMatrix.Columns.Item("212").Cells.Item(i).Specific.Value.ToString();// oSalesQuotation.Lines.UoMCode;
                                                detalle.Almacen = oMatrix.Columns.Item("24").Cells.Item(i).Specific.Value.ToString();// oSalesQuotation.Lines.WarehouseCode;
                                                detalle.Margen = oMatrix.Columns.Item("2004").Cells.Item(i).Specific.Value.ToString();// oSalesQuotation.Lines.CostingCode;
                                                detalle.CentroCosto = oMatrix.Columns.Item("2003").Cells.Item(i).Specific.Value.ToString();// oSalesQuotation.Lines.CostingCode2;
                                                detalle.CentroBeneficio = oMatrix.Columns.Item("2002").Cells.Item(i).Specific.Value.ToString();// oSalesQuotation.Lines.CostingCode3;
                                                detalle.Equipos = oMatrix.Columns.Item("2001").Cells.Item(i).Specific.Value.ToString();// oSalesQuotation.Lines.CostingCode4;
                                                detalle.CC5 = oMatrix.Columns.Item("2000").Cells.Item(i).Specific.Value.ToString();// oSalesQuotation.Lines.CostingCode5;
                                                detalle.Proyecto = oMatrix.Columns.Item("31").Cells.Item(i).Specific.Value.ToString();// oSalesQuotation.Lines.ProjectCode;
                                                detalle.GrupoDetraccion = oMatrix.Columns.Item("U_EXX_GRUPODET").Cells.Item(i).Specific.Value.ToString();// oSalesQuotation.Lines.UserFields.Fields.Item("U_EXX_GRUPODET").Value;
                                                detalle.GrupoRetencion = oMatrix.Columns.Item("U_EXX_GRUPOPER").Cells.Item(i).Specific.Value.ToString();// oSalesQuotation.Lines.UserFields.Fields.Item("U_EXX_GRUPOPER").Value;
                                                detalle.NroOferta = iDocEntry.ToString();
                                                detalle.LineaOferta = iLinea.ToString();

                                                Doc.Lineas.Add(detalle);
                                            }
                                            else
                                            {
                                                Globals.SBO_Application.SetStatusBarMessage("La linea seleccionada con el Item " + ItemCode + " ya fue asignado a otra Solicitud de Compra", SAPbouiCOM.BoMessageTime.bmt_Short, true);

                                            }

                                        }



                                    }
                                }
                            }
                            Globals.Release(oSalesQuotation);
                            PurchaseRequest(Doc);
                        }
                        else
                        {
                            Globals.SBO_Application.SetStatusBarMessage("El documento no esta aprobado", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                        }
                    }
                    else
                    {
                        Globals.Release(oSalesQuotation);
                    }
                }
                else
                {
                    Globals.SBO_Application.SetStatusBarMessage("Debe seleccionar un Documento", SAPbouiCOM.BoMessageTime.bmt_Short, false);

                }

            }
            catch (Exception ex)
            {
                throw ex;
            }


        }

        public static void PurchaseRequest(OfertaVentaDatos ofertaVenta)
        {
            try
            {
                Globals.SBO_Application.ActivateMenuItem("39724");

                //SAPbouiCOM.Form oForm = Globals.SBO_Application.OpenForm((BoFormObjectEnum)1470000113, "", "");



                SAPbouiCOM.Form oForm = Globals.SBO_Application.Forms.ActiveForm;

                int iFormType = -1 * oForm.Type;
                SAPbouiCOM.Form oUserForm = Globals.SBO_Application.Forms.GetFormByTypeAndCount(iFormType, oForm.TypeCount);


                SAPbouiCOM.ComboBox oComboSuc = oForm.Items.Item("2001").Specific;
                oComboSuc.Select(ofertaVenta.Sucursal.ToString());
                //oComboSuc.Select("1");

                try
                {
                    oForm.Items.Item("U_vOt").Specific.Value = ofertaVenta.OT;
                }
                catch (Exception ex)
                {

                    oUserForm.Items.Item("U_vOt").Specific.Value = ofertaVenta.OT;
                }

                try
                {
                    oForm.Items.Item("U_vTallCode").Specific.Value = ofertaVenta.Taller;
                }
                catch (Exception ex)
                {

                    oUserForm.Items.Item("U_vTallCode").Specific.Value = ofertaVenta.Taller;
                }

                SAPbouiCOM.Matrix oMatrix = Globals.SBO_Application.Forms.ActiveForm.Items.Item("38").Specific;
                int r = oMatrix.RowCount;
                foreach (DetalleOV linea in ofertaVenta.Lineas)
                {
                    oMatrix.Columns.Item("1").Cells.Item(r).Specific.Value = linea.ItemCode;
                    oMatrix.Columns.Item("3").Cells.Item(r).Specific.Value = linea.Descripcion;
                    oMatrix.Columns.Item("24").Cells.Item(r).Specific.Value = linea.Almacen;
                    oMatrix.Columns.Item("11").Cells.Item(r).Specific.Value = linea.Cantidad;
                    oMatrix.Columns.Item("U_EXS_CU").Cells.Item(r).Specific.Value = linea.Precio;
                    oMatrix.Columns.Item("160").Cells.Item(r).Specific.Value = linea.TaxCode;

                    if(linea.Proyecto != "")
                        oMatrix.Columns.Item("31").Cells.Item(r).Specific.Value = linea.Proyecto;
                    if (linea.Margen != "")
                        oMatrix.Columns.Item("2004").Cells.Item(r).Specific.Value = linea.Margen;
                    if (linea.CentroCosto != "")
                        oMatrix.Columns.Item("2003").Cells.Item(r).Specific.Value = linea.CentroCosto;
                    if (linea.CentroBeneficio != "")
                        oMatrix.Columns.Item("2002").Cells.Item(r).Specific.Value = linea.CentroBeneficio;
                    if (linea.Margen != "")
                        oMatrix.Columns.Item("2001").Cells.Item(r).Specific.Value = linea.Equipos;
                    if (linea.Equipos != "")
                        oMatrix.Columns.Item("2000").Cells.Item(r).Specific.Value = linea.CC5;

                    //oMatrix.Columns.Item("U_EXX_GRUPODET").Cells.Item(r).Specific.Value = linea.GrupoDetraccion;

                    if(linea.GrupoDetraccion  != "")
                    {
                        SAPbouiCOM.ComboBox oCombo = oMatrix.Columns.Item("U_EXX_GRUPODET").Cells.Item(r).Specific;
                        oCombo.Select(linea.GrupoDetraccion);
                    }
                    if (linea.GrupoRetencion != "")
                    {
                        SAPbouiCOM.ComboBox oCombo = oMatrix.Columns.Item("U_EXX_GRUPOPER").Cells.Item(r).Specific;
                        oCombo.Select(linea.GrupoRetencion);
                    }
                    oMatrix.Columns.Item("U_EXX_OFBA").Cells.Item(r).Specific.Value = linea.NroOferta;
                    oMatrix.Columns.Item("U_EXX_LOFB").Cells.Item(r).Specific.Value = linea.LineaOferta;

                    //oMatrix.AddRow();
                    r++;
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
