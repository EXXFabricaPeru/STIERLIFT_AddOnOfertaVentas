using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using AddOnUI.App;

namespace AddOnUI.Logica
{
    public class FacturaCompra
    {
        public static void GenerarControles(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.Item oItem = oForm.Items.Add("BtnAnu", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 258;
                oItem.Top = 540;
                oItem.Width = 150;
                oItem.Height = 20;

                SAPbouiCOM.Button oBtn = ((SAPbouiCOM.Button)(oItem.Specific));
                oBtn.Caption = "Anular Compra";
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void AnularCompra(SAPbouiCOM.Form oForm)
        {
            try
            {
                int iValor = Globals.SBO_Application.MessageBox("Desea anular la factura?", 3, "Si", "No", "Quizas");

                if (iValor == 1)
                {
                    SAPbouiCOM.EditText oDocNum = oForm.Items.Item("8").Specific;

                    Globals.Query = "SELECT \"DocEntry\" FROM OPCH WHERE \"DocNum\"  ='" + oDocNum.Value.ToString() + "' ";
                    Globals.RunQuery(Globals.Query);
                    int iDocEntry = Int32.Parse(Globals.oRec.Fields.Item(0).Value.ToString());
                    Globals.Release(Globals.oRec);

                    //anular
                    SAPbobsCOM.Documents oDocSAP = (SAPbobsCOM.Documents)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);

                    if (oDocSAP.GetByKey(iDocEntry))
                    {
                        SAPbobsCOM.Documents oDocCancel = oDocSAP.CreateCancellationDocument();
                        //modificar campos en el documento de cancelacion
                        //--
                        //--
                        Globals.lRetCode = oDocCancel.Add();

                        Globals.Release(oDocSAP);
                        Globals.Release(oDocCancel);

                        if (Globals.lRetCode == 0)
                        {
                            Globals.SBO_Application.SetStatusBarMessage("Documento anulado correctamente", SAPbouiCOM.BoMessageTime.bmt_Short, false);

                            Globals.SBO_Application.ActivateMenuItem("1304");
                        }
                        else
                        {
                            Globals.oCompany.GetLastError(out Globals.lErrCode, out Globals.sErrMsg);
                            throw new Exception(String.Concat(Globals.lErrCode, ": ", Globals.sErrMsg));
                        }
                    }
                    else
                    {
                        Globals.Release(oDocSAP);
                        Globals.oCompany.GetLastError(out Globals.lErrCode, out Globals.sErrMsg);
                        throw new Exception(String.Concat(Globals.lErrCode, ": ", Globals.sErrMsg));
                    }



                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
