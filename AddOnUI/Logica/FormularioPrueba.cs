using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using AddOnUI.App;

namespace AddOnUI.Logica
{
    public class FormularioPrueba
    {
        public static void CargarFormulario()
        {
            SAPbouiCOM.Form oForm = default(SAPbouiCOM.Form);

            try
            {
                oForm = Globals.SBO_Application.Forms.Item("PruebasForm_2");
                Globals.SBO_Application.MessageBox("El formulario ya se encuentra abierto");
            }
            catch
            {
                SAPbouiCOM.FormCreationParams fcp = default(SAPbouiCOM.FormCreationParams);
                fcp = Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;
                fcp.FormType = "PruebasForm";
                fcp.UniqueID = "PruebasForm_2";
                string FormName = "\\PruebasForm.srf";
                fcp.XmlData = Globals.LoadFromXML(ref FormName);
                oForm = Globals.SBO_Application.Forms.AddEx(fcp);
            }
            oForm.Top = 50;
            oForm.Left = 200;
            oForm.Visible = true;
        }

        public static void Cerrar(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Close();
                return;
            }
            catch (Exception ex)
            {
                throw ex;
            } 
        }

        public static void Cargar(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.EditText oFechIni = oForm.Items.Item("Item_2").Specific;
                SAPbouiCOM.EditText oFechFin = oForm.Items.Item("Item_3").Specific;

                Globals.Query = "SELECT 'N' \"Marcar\",\"DocEntry\",\"DocNum\",\"DocEntry\" \"NroSAP\",\"CardCode\"  \n";
                Globals.Query += " FROM OINV \n";
                Globals.Query += " WHERE \"CANCELED\"='N' AND \"DocStatus\" <> 'C' \n";
                Globals.Query += " AND \"DocDate\" >= '"+oFechIni.Value.ToString()+"' \n";
                Globals.Query += " AND \"DocDate\" <= '" + oFechFin.Value.ToString() + "' \n";

                LoadMatrix(oForm, Globals.Query);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void LoadMatrix(SAPbouiCOM.Form oForm, string Query)
        {
            try
            {
                SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("Item_10").Specific;
                SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("DT_0");

                if (oDataTable.Rows.Count > 0)
                {
                    oDataTable.Clear();
                }

                oDataTable.ExecuteQuery(Query);

                oMatrix.Columns.Item("Col_0").DataBind.Bind("DT_0", "Marcar");
                oMatrix.Columns.Item("Col_1").DataBind.Bind("DT_0", "DocEntry");
                oMatrix.Columns.Item("Col_2").DataBind.Bind("DT_0", "DocNum");
                oMatrix.Columns.Item("Col_3").DataBind.Bind("DT_0", "NroSAP");
                oMatrix.Columns.Item("Col_4").DataBind.Bind("DT_0", "CardCode");

                oMatrix.Clear();
                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void AnularFV(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("Item_10").Specific;

                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    SAPbouiCOM.CheckBox oCheckBox = oMatrix.Columns.Item("Col_0").Cells.Item(i).Specific;
                    bool bVal = (oCheckBox.Checked) ? true : false;
                    int iDocEntry = Int32.Parse(oMatrix.Columns.Item("Col_1").Cells.Item(i).Specific.Value.ToString());

                    if (bVal)
                    {
                        //anular registro
                        SAPbobsCOM.Documents oDoc
                            = (SAPbobsCOM.Documents)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);

                        if (oDoc.GetByKey(iDocEntry))
                        {
                            SAPbobsCOM.Documents oDocCancel = oDoc.CreateCancellationDocument();

                            if (oDocCancel.Add().Equals(0))
                            {
                                Globals.SBO_Application.SetStatusBarMessage("Cancelación creada correctamente: " + iDocEntry.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, false);
                            }
                            else
                            {
                                Globals.oCompany.GetLastError(out Globals.lErrCode, out Globals.sErrMsg);

                                Globals.SBO_Application.SetStatusBarMessage("Cancelación incorrecta: " + iDocEntry.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            }
                            Globals.Release(oDocCancel);
                        }
                        else
                        {
                            Globals.SBO_Application.SetStatusBarMessage("Registro no existe en SAP: " + iDocEntry.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, false);
                        }
                        Globals.Release(oDoc);
                    }
                }
                Globals.SBO_Application.MessageBox("Proceso Culminado");

                Cargar(oForm);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void LinkPressed(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("Item_10").Specific;

                switch (pVal.ColUID)
                {
                    case "Col_3":
                        SAPbouiCOM.Column oColumn3 = oMatrix.Columns.Item("Col_3");
                        SAPbouiCOM.LinkedButton oLinkedButton3;
                        oLinkedButton3 = oColumn3.ExtendedObject;
                        oLinkedButton3.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Invoice;
                        break;
                    case "Col_4":
                        SAPbouiCOM.Column oColumn4 = oMatrix.Columns.Item("Col_4");
                        SAPbouiCOM.LinkedButton oLinkedButton4;
                        oLinkedButton4 = oColumn4.ExtendedObject;
                        oLinkedButton4.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner;
                        break;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
