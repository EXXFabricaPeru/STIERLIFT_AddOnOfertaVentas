using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using AddOnUI.App;

namespace AddOnUI.Logica
{
    public class Validaciones
    {
        public static void Validar(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, int FormType)
        {
            try
            {
                switch (FormType)
                {
                    case 133://Factura de venta
                        ValidarVentas(oForm, pVal);
                        break;

                    case 141://factura de compra
                        ValidarCompras(oForm, pVal);
                        break;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void ValidarVentas(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                int iFormType = -1 * pVal.FormType;

                SAPbouiCOM.Form oUserForm = Globals.SBO_Application.Forms.GetFormByTypeAndCount(iFormType, pVal.FormTypeCount);

                SAPbouiCOM.EditText oTipo = oUserForm.Items.Item("U_SBA_TIPOD").Specific;
                SAPbouiCOM.EditText oSerie = oUserForm.Items.Item("U_SBA_SERIED").Specific;
                SAPbouiCOM.EditText oCorr = oUserForm.Items.Item("U_SBA_CORRD").Specific;

                if (oTipo.Value.ToString().Trim() == "")
                {
                    throw new Exception("El tipo de documento es obligatorio");
                }
                if (oSerie.Value.ToString().Trim() == "")
                {
                    throw new Exception("La serie de documento es obligatorio");
                }
                if (oCorr.Value.ToString().Trim() == "")
                {
                    throw new Exception("El correlativo de documento es obligatorio");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void ValidarCompras(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                int iFormType = -1 * pVal.FormType;

                SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("38").Specific;
                SAPbouiCOM.Form oUserForm = Globals.SBO_Application.Forms.GetFormByTypeAndCount(iFormType, pVal.FormTypeCount);

                SAPbouiCOM.EditText oNroRef = oForm.Items.Item("14").Specific;


                if (oNroRef.Value.ToString().Trim() == "")
                {
                    throw new Exception("El campo nro de referencia es obligatorio");
                }

                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    if (oMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.ToString() != "" &&
                        oMatrix.Columns.Item("31").Cells.Item(i).Specific.Value.ToString() == "")
                    {
                        throw new Exception("El campo proyecto no puede quedar vacío. Validar línea " + i.ToString());
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
