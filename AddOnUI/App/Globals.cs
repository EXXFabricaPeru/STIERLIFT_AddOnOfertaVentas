using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;

namespace AddOnUI.App
{
    class Globals
    {
        public static SAPbobsCOM.Recordset oRec = default(SAPbobsCOM.Recordset);
        public static SAPbobsCOM.Recordset oRec2 = default(SAPbobsCOM.Recordset);
        
        public static int continuar = -1;
        public static SAPbobsCOM.Company oCompany;
        public static SAPbouiCOM.Application SBO_Application;

        public static SAPbobsCOM.CompanyService oCmpSrv;
        public static SAPbouiCOM.EventFilters oFilters;
        public static SAPbouiCOM.EventFilter oFilter;
        

        public static string Action = null;
        public static string Error = null;
        public static string ServerError = null;
        //public static int continuar = -1;

        public static int camposentotal = 0;
        public static int cuantosrepetidos = 0;

        public static string Addon = null;
        public static string version = null;
        public static int SAPVersion;
        public static string Query = null;
        public static string Query2 = null;
        public static string Query3 = null;
        public static string oldversion = "";
        public static bool Actual = false;

        public static int lErrCode = 0;
        public static int lRetCode;
        public static string sErrMsg = null;
        public static bool bError = false;
        public static string LogFile = null;

        public static string filename;
        private static string pathItem;
        public static SAPbouiCOM.Form DialogForm;
        public static Dictionary<int, List<String>> Diccionario = new Dictionary<int, List<String>>();
     
        public static bool BotonPresionado = false;

        public static String AddOnName = "AddOnMigrarOfV";
        public static String AddOnVersion = "1.0.0";
        public static Boolean BtnMigPres = false;

        #region MyRegion
        public static bool DocTypeServicio = false;
        public static bool TSC = false;
        public static bool TipoDocumento = false;
        public static String Tipo = null;
        public static String Serie = null;
        public static String Correlativo = null;
        //public static List<Recaudaciones.Model.Confba> objLstConfba;
        public static Double Retencion = 700;

        public static String DocEntryOV = "";
        public static String DocEntryFAC = "";
        public static String DocEntryPedido = "";
        public static String DocEntryRequerimiento = "";

        public static String TipoSalida = "00";


        public static string GetUDFValue(string Field, string Type, int FormType, int FormTypeCount)
        {
            FormType = -1 * FormType;
            SAPbouiCOM.Form oUserForm = Globals.SBO_Application.Forms.GetFormByTypeAndCount(FormType, FormTypeCount);
            if (Type == "ComboBox")
            {
                SAPbouiCOM.ComboBox oUDF = oUserForm.Items.Item(Field).Specific;
                return (oUDF.Value.Replace(" ", "") == "" | oUDF.Value == null) ? null : oUDF.Value;
            }
            else if (Type == "EditText")
            {
                SAPbouiCOM.EditText oUDF = oUserForm.Items.Item(Field).Specific;
                return (oUDF.Value.Replace(" ", "") == "" | oUDF.Value == null) ? null : oUDF.Value;
            }
            else
            {
                return null;
            }
        }

        //ObtenerMonedaMontoTC(string Importe, int p, string Fecha, ref string Moneda, ref string Importe_2, ref double TipoCambio)
        public static void ObtenerMonedaMontoTC(string ImporteTexto, int tipo, string sFecha, ref string Moneda, ref double Importe, ref double TipoCambio)
        {
            double ImporteConvertido = 0;
            if (tipo == 1)
            {
                string[] tokens = ImporteTexto.Split('(');
                ImporteTexto = tokens[1];
                string[] tokens2 = ImporteTexto.Split(')');
                ImporteConvertido = Convert.ToDouble(tokens2[0]);
                bool result = Double.TryParse(tokens2[0], out ImporteConvertido);
                if (!(result))
                {
                    throw new Exception("SYP(BPS) : El monto a convertir no tiene el formato correcto: " + tokens2[0]);
                }
            }
            else if (tipo == 2)
            {
                string[] tokens = ImporteTexto.Split(' ');
                bool result, result2 = false;
                ImporteTexto = tokens[1];
                result = Double.TryParse(tokens[0], out ImporteConvertido);
                if (!result)
                {
                    result2 = Double.TryParse(tokens[2], out ImporteConvertido);
                }
                if (!(result) && !(result2))
                {
                    throw new Exception("SYP(BPS) : El monto a convertir no tiene el formato correcto: " + tokens[0]);
                }
            }
            else if (tipo == 3)
            {
                string[] tokens = ImporteTexto.Split(' ');
                bool result, result2 = false;
                double TC = 0;
                ImporteTexto = tokens[1];
                result = Double.TryParse(tokens[0], out ImporteConvertido);
                if (!result)
                {
                    result2 = Double.TryParse(tokens[1], out ImporteConvertido);
                    //if (tokens[0] != "S/")
                    //{
                    //    ClassGC.ObtenerTC(sFecha, tokens[0], ref TC);
                    //}
                    Moneda = tokens[0];
                    Importe = ImporteConvertido;
                    TipoCambio = 1;
                }
                if (!(result) && !(result2))
                {
                    throw new Exception("SYP(BPS) : El monto a convertir no tiene el formato correcto: " + tokens[0]);
                }
            }
        }


        #endregion

        public static bool IsHana()
        {
            try
            {
                if (Globals.oCompany.DbServerType == (SAPbobsCOM.BoDataServerTypes)9)
                    return true;
                else
                    return false;
            }
            catch (Exception ex)
            {
                Globals.SBO_Application.MessageBox(ex.Message);
                return false;
            }
        }

        public static void WriteLogTxt(string x, string filename)
        {
            //System.Windows.Forms.Application.StartupPath;
            string path = System.Windows.Forms.Application.StartupPath + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string FILE_NAME = path + "\\" + filename + ".txt";
            if (System.IO.File.Exists(FILE_NAME) == false)
            {
                System.IO.File.Create(FILE_NAME).Dispose();
            }
            System.IO.StreamWriter objWriter = new System.IO.StreamWriter(FILE_NAME, true, Encoding.Default);
            objWriter.WriteLine(x);
            objWriter.Close();
        }

        public static void StartTransaction()
        {
            try
            {
                oCompany.StartTransaction();
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public static void CommitTransaction()
        {
            try
            {
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
            }
            catch (Exception ex)
            {
                //throw new Exception(ex.Message);
                throw ex;
            }
        }
        public static void RollBackTransaction()
        {
            try
            {
                if (oCompany.InTransaction)
                {
                    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        //public static string Query = null;
        public static SAPbobsCOM.Recordset RunQuery(string Query)
        {
            try
            {
                oRec = Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRec.DoQuery(Query);
                return oRec;
            }
            catch (Exception ex)
            {
                Globals.SBO_Application.MessageBox(ex.Message);
                return null;
            }
        }
        public static object Release(object objeto)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objeto);
            Query = null;
            GC.Collect();
            return null;
        }

        //public static string Query2 = null;
        public static SAPbobsCOM.Recordset RunQuery2(string Query)
        {
            try
            {
                oRec2 = Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRec2.DoQuery(Query);
                return oRec2;
            }
            catch (Exception ex)
            {
                Globals.SBO_Application.MessageBox(ex.Message);
                return null;
            }
        }
        public static object Release2(object objeto)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objeto);
            Query = null;
            GC.Collect();
            return null;
        }

        public static string LoadFromXML(ref string FileName)
        {
            System.Xml.XmlDocument oXmlDoc = null;
            string sPath = null;
            oXmlDoc = new System.Xml.XmlDocument();
            sPath = System.Windows.Forms.Application.StartupPath;
            oXmlDoc.Load(sPath + FileName);
            return (oXmlDoc.InnerXml);
        }

        public static void OpenFile(SAPbouiCOM.Form oForm, string path)
        {
            try
            {
                pathItem = path;
                DialogForm = oForm;
                System.Threading.Thread ShowFolderBrowserThread;
                ShowFolderBrowserThread = new System.Threading.Thread(ShowFolderBrowser);
                if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Unstarted)
                {
                    ShowFolderBrowserThread.SetApartmentState(ApartmentState.STA);
                    ShowFolderBrowserThread.Start();
                }
                else if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Stopped)
                {
                    ShowFolderBrowserThread.Start();
                    ShowFolderBrowserThread.Join();
                }
                //ShowFolderBrowserThread.Abort();
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }
        public static void ShowFolderBrowser()
        {
            try
            {
                NativeWindow nws = new NativeWindow();
                OpenFileDialog MyTest = new OpenFileDialog();
                MyTest.Multiselect = false;
                MyTest.Filter = "Text Files (.txt)|*.txt";
                Process[] MyProcs = null;
                //string filename = null;
                MyProcs = Process.GetProcessesByName("SAP Business One");
                nws.AssignHandle(System.Diagnostics.Process.GetProcessesByName("SAP Business One")[0].MainWindowHandle);
                if (MyTest.ShowDialog(nws) == System.Windows.Forms.DialogResult.OK)
                {
                    filename = MyTest.FileName;
                    DialogForm.Items.Item(pathItem).Specific.Value = filename;
                    DialogForm = null;
                    pathItem = null;
                }
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }
        public static double ConvertDouble(string ImporteTexto, int tipo)
        {
            //Console.WriteLine(ImporteTexto);
            double ImporteConvertido = 0;
            if (tipo == 1)
            {
                string[] tokens = ImporteTexto.Split('(');
                ImporteTexto = tokens[1];
                string[] tokens2 = ImporteTexto.Split(')');
                ImporteConvertido = Convert.ToDouble(tokens2[0]);
                bool result = Double.TryParse(tokens2[0], out ImporteConvertido);
                if (!(result))
                {
                    throw new Exception("SYP(BPS) : El monto a convertir no tiene el formato correcto: " + tokens2[0]);
                }
            }
            else if (tipo == 2)
            {
                string[] tokens = ImporteTexto.Split(' ');
                bool result, result2 = false;
                ImporteTexto = tokens[1];
                result = Double.TryParse(tokens[0], out ImporteConvertido);
                if (!result)
                {
                    result2 = Double.TryParse(tokens[2], out ImporteConvertido);
                }
                if (!(result) && !(result2))
                {
                    throw new Exception("SYP(BPS) : El monto a convertir no tiene el formato correcto: " + tokens[0]);
                }
            }
            else if (tipo == 3)
            {
                string[] tokens = ImporteTexto.Split(' ');
                bool result, result2 = false;
                result = Double.TryParse(tokens[0], out ImporteConvertido);
                if (!result)
                {
                    result2 = Double.TryParse(tokens[1], out ImporteConvertido);
                }
                if (!(result) && !(result2))
                {
                    throw new Exception("SYP(BPS) : El monto a convertir no tiene el formato correcto: " + tokens[0]);
                }
            }

            return ImporteConvertido;
        }
       

        internal static void AgregarMenu()
        {
            SAPbouiCOM.Menus oMenus = default(SAPbouiCOM.Menus);
            SAPbouiCOM.MenuItem oMenuItem = default(SAPbouiCOM.MenuItem);
            oMenus = Globals.SBO_Application.Menus;
            SAPbouiCOM.MenuCreationParams oCreationPackage = default(SAPbouiCOM.MenuCreationParams);
            oCreationPackage = Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
            int cont = 0;
            Globals.Query = "select count(1) as Existe from \"OUDO\" where \"TableName\" = 'SYP_EQUI_SERV'";
            Globals.RunQuery(Globals.Query);
            cont = Globals.oRec.Fields.Item(0).Value;
            Globals.Release(Globals.oRec);
            //if (cont == 0)
            //{
            try
            {
                oMenuItem = Globals.SBO_Application.Menus.Item("8192");
                oMenus = oMenuItem.SubMenus;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "UDO_SYP_RECAUDACION";
                oCreationPackage.String = "Crear UDO Recaudación";
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception ex)
            {
                Globals.SBO_Application.SetStatusBarMessage(ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, false);
            }
            //}
        }
    }
}
