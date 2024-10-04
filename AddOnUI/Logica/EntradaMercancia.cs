using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using AddOnUI.App;

namespace AddOnUI.Logica
{
    public class EntradaMercancia
    {
        public static void GenerarControles(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.Item oItem = oForm.Items.Add("BtnAnu", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 258;
                oItem.Top = 100;
                oItem.Width = 150;
                oItem.Height = 20;

                SAPbouiCOM.Button oBtn = ((SAPbouiCOM.Button)(oItem.Specific));
                oBtn.Caption = "Anular Entrada";
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void AnularEntrada(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.EditText eDocNum = oForm.Items.Item("7").Specific;

                Globals.Query = "SELECT \"DocEntry\" FROM OIGN WHERE \"DocNum\"= " + eDocNum.Value.ToString();
                Globals.RunQuery(Globals.Query);
                int iDocEntry = Int32.Parse(Globals.oRec.Fields.Item(0).Value.ToString());
                Globals.Release(Globals.oRec);

                Globals.Query = "SELECT COUNT(1) FROM OIGE WHERE U_SBA_CONTROLC = '" + iDocEntry.ToString() + "'";
                Globals.RunQuery(Globals.Query);
                int iCont = Int32.Parse(Globals.oRec.Fields.Item(0).Value.ToString());
                Globals.Release(Globals.oRec);

                if (iCont > 0)
                {
                    throw new Exception("La entrada de mercancia " + eDocNum.Value.ToString() + " ya tiene una salida asociada.");
                }
                else
                {
                    SAPbouiCOM.EditText eDocDate = oForm.Items.Item("9").Specific;
                    SAPbouiCOM.EditText eTaxDate = oForm.Items.Item("38").Specific;
                    string sDocDate = eDocDate.Value.ToString();
                    string sTaxDate = eTaxDate.Value.ToString();

                    SAPbobsCOM.Documents oDocSAP =
                        ( SAPbobsCOM.Documents)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
                    oDocSAP.DocDate = DateTime.Parse(sDocDate.Substring(6, 2) + "/" + sDocDate.Substring(4, 2) + "/" + sDocDate.Substring(0, 4));
                    oDocSAP.TaxDate = DateTime.Parse(sTaxDate.Substring(6, 2) + "/" + sTaxDate.Substring(4, 2) + "/" + sTaxDate.Substring(0, 4));
                    oDocSAP.Comments = "Created by SDK";
                    oDocSAP.UserFields.Fields.Item("U_SBA_CONTROLC").Value = iDocEntry.ToString();

                    Globals.Query = "SELECT \"ItemCode\",\"Quantity\",\"WhsCode\",\"AcctCode\" FROM IGN1 WHERE \"DocEntry\"= " + iDocEntry.ToString();
                    Globals.RunQuery(Globals.Query);
                    while (!Globals.oRec.EoF)
                    {
                        oDocSAP.Lines.ItemCode = Globals.oRec.Fields.Item(0).Value.ToString();
                        oDocSAP.Lines.Quantity = double.Parse(Globals.oRec.Fields.Item(1).Value.ToString());
                        oDocSAP.Lines.WarehouseCode = Globals.oRec.Fields.Item(2).Value.ToString();
                        oDocSAP.Lines.AccountCode = Globals.oRec.Fields.Item(3).Value.ToString();

                        oDocSAP.Lines.Add();

                        Globals.oRec.MoveNext();
                    }
                    Globals.Release(Globals.oRec);

                    if (!oDocSAP.Add().Equals(0))
                    {
                        Globals.oCompany.GetLastError(out Globals.lErrCode, out Globals.sErrMsg);
                        Globals.Release(oDocSAP);
                        throw new Exception(String.Concat(Globals.lErrCode.ToString(), ": ", Globals.sErrMsg));
                    }
                    else
                    {
                        string sValor = Globals.oCompany.GetNewObjectKey();
                        Globals.Release(oDocSAP);
                        Globals.SBO_Application.SetStatusBarMessage("Proceso culminado correctamente. Salida Nro: " + sValor, SAPbouiCOM.BoMessageTime.bmt_Short, false);
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
