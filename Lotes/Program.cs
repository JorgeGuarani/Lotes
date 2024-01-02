using System;
using System.Collections.Generic;
using SAPbouiCOM.Framework;
using SAPbouiCOM;
using System.Windows.Forms;

namespace Lotes
{
    
    class Program
    {
        public static string v_LoteNom = "";
        public static int v_num = 0;
        public static SAPbouiCOM.EditText v_codigo = null;
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                SAPbouiCOM.Framework.Application oApp = null;
                if (args.Length < 1)
                {
                    oApp = new SAPbouiCOM.Framework.Application();
                }
                else
                {
                    oApp = new SAPbouiCOM.Framework.Application(args[0]);
                }
                Menu MyMenu = new Menu();
                MyMenu.AddMenuItems();
                oApp.RegisterMenuEventHandler(MyMenu.SBO_Application_MenuEvent);
                SAPbouiCOM.Framework.Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                SAPbouiCOM.Framework.Application.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                SAPbouiCOM.Framework.Application.SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private static void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                #region ENTRADA DE MERCADERIA - LOTE             
                if (pVal.FormTypeEx == "41")
                {
                    //agarramos el formulario
                    SAPbouiCOM.Form form = SAPbouiCOM.Framework.Application.SBO_Application.Forms.Item(FormUID);
                    

                    if(pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.BeforeAction == false){
                        //consultamos la descripcion
                        SAPbobsCOM.Recordset oNumLote;
                        oNumLote = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oNumLote.DoQuery("SELECT \"Name\",\"U_LoteNum\" FROM \"@LOTENUM\" ");
                        v_LoteNom = oNumLote.Fields.Item(0).Value.ToString();   
                        v_num = int.Parse(oNumLote.Fields.Item(1).Value.ToString());
                    }
                    //evento al presionar tabulador
                    #region TABULADOR
                    if (pVal.EventType == BoEventTypes.et_LOST_FOCUS && pVal.BeforeAction==false && pVal.ItemUID == "3" && pVal.ColUID=="2") 
                    {
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)form.Items.Item("3").Specific;
                        SAPbouiCOM.EditText v_texto = (SAPbouiCOM.EditText)oMatrix.Columns.Item("2").Cells.Item(pVal.Row).Specific;
                        string texto_v = v_texto.Value;            
                        if((string.IsNullOrEmpty(texto_v)))//if (pVal.CharPressed == (char)9 && pVal.ColUID=="2" )
                        {
                            SAPbouiCOM.EditText oLote = (SAPbouiCOM.EditText)oMatrix.Columns.Item("2").Cells.Item(pVal.Row).Specific;
                            v_num = v_num + 1;
                            oLote.Value = v_LoteNom + v_num;
                            oMatrix.Columns.Item("3").Cells.Item(pVal.Row).Click();
                            //if (pVal.Row > 1)
                            //{
                            //    SAPbouiCOM.EditText oLoteNumer = (SAPbouiCOM.EditText)oMatrix.Columns.Item("2").Cells.Item(pVal.Row-1).Specific;
                            //    string v_numeroLote = oLoteNumer.Value;
                            //    v_numeroLote = v_numeroLote.Replace(v_LoteNom, string.Empty);
                            //    int v_nuevoLote = v_num + 1;//int.Parse(v_numeroLote) + 1;
                            //    v_numeroLote = v_LoteNom + v_nuevoLote.ToString();

                            //    SAPbouiCOM.EditText oLoteNuevo = (SAPbouiCOM.EditText)oMatrix.Columns.Item("2").Cells.Item(pVal.Row).Specific;
                            //    oLoteNuevo.Value = v_numeroLote;
                            //    oMatrix.Columns.Item("3").Cells.Item(pVal.Row).Click();
                            //}
                            //else
                            //{
                            //    //agarramos los datos del formulario                           
                            //    SAPbouiCOM.EditText oLote = (SAPbouiCOM.EditText)oMatrix.Columns.Item("2").Cells.Item(pVal.Row).Specific;
                            //    SAPbouiCOM.EditText oCan = (SAPbouiCOM.EditText)oMatrix.Columns.Item("3").Cells.Item(pVal.Row).Specific;
                            //    //realizamos una consulta a la tabla
                            //    SAPbobsCOM.Recordset oNumLote1;
                            //    oNumLote1 = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            //    oNumLote1.DoQuery("SELECT \"U_LoteNum\"+1 FROM \"@LOTENUM\" ");
                            //    string v_LoteNum = oNumLote1.Fields.Item(0).Value.ToString();
                            //    //mandamos el numero de lote a la tabla
                            //    oLote.Value = v_LoteNom + v_num;
                            //    oMatrix.Columns.Item("3").Cells.Item(pVal.Row).Click();
                            //}
                           
                        }
                    }
                    #endregion
                }
                #endregion

                #region FORM ENTRADA DE MERCADERIA
                if (pVal.FormTypeEx == "721")
                {
                    //agarramos el formulario
                    SAPbouiCOM.Form form = SAPbouiCOM.Framework.Application.SBO_Application.Forms.Item(FormUID);

                    //evento antes de crear el documento
                    if (pVal.ItemUID=="1" && pVal.BeforeAction==true && pVal.EventType==BoEventTypes.et_ITEM_PRESSED)
                    {
                        v_codigo = (SAPbouiCOM.EditText)form.Items.Item("8").Specific;
                    }

                    //evento después de crear el documento
                    if(pVal.ItemUID=="1" && pVal.FormMode==3 && pVal.BeforeAction==false && pVal.ActionSuccess && pVal.EventType==BoEventTypes.et_ITEM_PRESSED)
                    {
                        //realizamos una consulta a la tabla
                        //SAPbobsCOM.Recordset oNumLote1;
                        //oNumLote1 = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        //oNumLote1.DoQuery("SELECT \"U_LoteNum\"+1, \"Name\" FROM \"@LOTENUM\" ");
                        //string v_LoteNum = oNumLote1.Fields.Item(0).Value.ToString();
                        //int v_lote = int.Parse(v_LoteNum);
                        //string v_nombre = oNumLote1.Fields.Item(1).Value.ToString();

                        ////consultamos los datos a actualizar
                        //SAPbobsCOM.Recordset oDatos;
                        //oDatos = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        //oDatos.DoQuery("SELECT \"ItemCode\",\"AbsEntry\",\"DistNumber\" FROM OBTN WHERE \"DistNumber\" Like 'LT%' ORDER BY \"AbsEntry\" ASC");
                        ////recorremos
                        //while (!oDatos.EoF)
                        //{
                        //    string v_entry = oDatos.Fields.Item(1).Value.ToString();
                        //    //el nombre del nuevo lote
                        //    string v_lotenew = v_nombre + v_lote.ToString();
                        //    //actualizamos los lotes nuevos
                        //    SAPbobsCOM.Recordset oUpdate;
                        //    oUpdate = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        //    oUpdate.DoQuery("UPDATE OBTN SET \"DistNumber\"='"+v_lotenew+ "' WHERE \"AbsEntry\"='"+v_entry+"' ");
                        //    oDatos.MoveNext();
                        //    v_lote++;
                        //    v_lotenew = "";
                        //}
                        //actualizamos la tabla de lote
                        //v_lote = v_lote - 1;
                        SAPbobsCOM.Recordset oUpdateLote;
                        oUpdateLote = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oUpdateLote.DoQuery("UPDATE \"@LOTENUM\" SET \"U_LoteNum\"='" + v_num + "' WHERE \"Code\"='1' ");
                        v_num = 0;

                    }
                }
                #endregion

                #region FORM COMPRAS ENTRADA DE MERCANCIAS
                if (pVal.FormTypeEx == "143")
                {
                    //agarramos el formulario
                    SAPbouiCOM.Form form = SAPbouiCOM.Framework.Application.SBO_Application.Forms.Item(FormUID);

                    //evento antes de crear el documento
                    if (pVal.ItemUID == "1" && pVal.BeforeAction == true && pVal.EventType == BoEventTypes.et_ITEM_PRESSED)
                    {
                        v_codigo = (SAPbouiCOM.EditText)form.Items.Item("8").Specific;
                    }

                    //evento después de crear el documento
                    if (pVal.ItemUID == "1" && pVal.FormMode == 3 && pVal.BeforeAction == false && pVal.ActionSuccess && pVal.EventType == BoEventTypes.et_ITEM_PRESSED)
                    {
                        SAPbobsCOM.Recordset oUpdateLote;
                        oUpdateLote = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oUpdateLote.DoQuery("UPDATE \"@LOTENUM\" SET \"U_LoteNum\"='" + v_num + "' WHERE \"Code\"='1' ");
                        v_num = 0;
                    }
                }
                #endregion

            }
            catch (Exception e)
            {

            }
        }

        public static void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.BeforeAction==false && (pVal.MenuUID == "1293") )
            {
                //cuando se elimina una fila del form entreda de mercaderias - lote
                renumerar();
            }
        }

        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }

        private static void renumerar()
        {
            try
            {
                SAPbouiCOM.Form form = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
                //consultamos la descripcion
                SAPbobsCOM.Recordset oNumLote;
                oNumLote = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oNumLote.DoQuery("SELECT \"Name\" FROM \"@LOTENUM\" ");
                string v_LoteNom = "LT";// oNumLote.Fields.Item(0).Value.ToString();

                string v_caption = form.Title.ToString();
                //verificamos si es el form de entrada de mercaderias
                if (v_caption.Equals("Lotes: Definir"))
                {
                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)form.Items.Item("3").Specific;
                    SAPbouiCOM.EditText oCantLotes = (SAPbouiCOM.EditText)form.Items.Item("6").Specific;
                    string cantLote = oCantLotes.Value;
                    cantLote = cantLote.Replace(",000",string.Empty);
                    int v_cant = oMatrix.RowCount;
                    int v_fila = 1;
                    if (v_cant > 1)
                    {
                        while (v_fila <= int.Parse(cantLote))
                        {
                            if (v_fila > 1)
                            {
                                SAPbouiCOM.EditText oLoteNumer = (SAPbouiCOM.EditText)oMatrix.Columns.Item("2").Cells.Item(v_fila - 1).Specific;
                                string v_numeroLote = oLoteNumer.Value;
                                v_numeroLote = v_numeroLote.Replace(v_LoteNom, string.Empty);
                                int v_nuevoLote = int.Parse(v_numeroLote) + 1;
                                v_numeroLote = v_LoteNom + v_nuevoLote.ToString();

                                SAPbouiCOM.EditText oLoteNuevo = (SAPbouiCOM.EditText)oMatrix.Columns.Item("2").Cells.Item(v_fila).Specific;
                                oLoteNuevo.Value = v_numeroLote;
                                oMatrix.Columns.Item("3").Cells.Item(v_fila).Click();
                            }
                            else
                            {
                                //agarramos los datos del formulario                           
                                SAPbouiCOM.EditText oLote = (SAPbouiCOM.EditText)oMatrix.Columns.Item("2").Cells.Item(v_fila).Specific;
                                SAPbouiCOM.EditText oCan = (SAPbouiCOM.EditText)oMatrix.Columns.Item("3").Cells.Item(v_fila).Specific;
                                //mandamos el numero de lote a la tabla
                                //realizamos una consulta a la tabla
                                SAPbobsCOM.Recordset oNumLote1;
                                oNumLote1 = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                oNumLote1.DoQuery("SELECT \"U_LoteNum\"+1 FROM \"@LOTENUM\" ");
                                string v_LoteNum = oNumLote1.Fields.Item(0).Value.ToString();

                                oLote.Value = v_LoteNom + v_LoteNum;
                                oMatrix.Columns.Item("3").Cells.Item(v_fila).Click();
                            }
                            v_fila++;
                        }
                    }
               }
            }
            catch(Exception e)
            {

            }
            
        }
    }
}
