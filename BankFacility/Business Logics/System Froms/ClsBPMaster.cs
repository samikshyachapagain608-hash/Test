using System;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using Microsoft.ApplicationBlocks.Data;
using System.Text;
using System.Xml;
using System.Collections.Generic;
using System.Collections;
using System.Threading;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Globalization;
using Sap.Data.Hana;

namespace BankFacility
{
    class ClsBPMaster:Base
    {

        #region VARIABLE DECLARATION
        SAPbouiCOM.Item _ExistItem, _NewItem, _Items;
        SAPbouiCOM.Button _Button;
        private bool flag;
        public SAPbouiCOM.Matrix _Matrix;
        SAPbouiCOM.StaticText _StaticText;
        string _DocEntry = "";
        CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep;
        string ReportPath = string.Empty;
        private static HanaConnection _conn;
        private static HanaDataAdapter _da;
        private static DataTable _dataTable;
        private static HanaParameter _parm1;
        private static HanaParameter _parm2;
        private static int _errors = 0;
        string strcon, _Sql;
        private SAPbouiCOM.EditText _EditText;
        private SAPbouiCOM.ComboBox _ComboBox;
        private SAPbouiCOM.DBDataSource _MasterDBDataSource, _ChildDBDataSource;
        private SAPbobsCOM.Recordset _RS;
        private int RowID { get; set; }
        private string _Code { get; set; }
        private string _LineId { get; set; }
        private string _ActivityType = string.Empty;
        string _BPCode = string.Empty, _BPName = string.Empty;

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR
        public ClsBPMaster()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsBPMaster()
        {
        }
        #endregion

        #region Item Event
        public override void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool oBubbleEvent)
        {
            #region Before Action
            if (pVal.Before_Action)
            {
                try
                {
                    if (pVal.ItemUID == "BtnBG")
                    {
                        _BPCode = _Form.DataSources.DBDataSources.Item("OCRD").GetValue("CardCode", 0).ToString();
                       
                        SAPbobsCOM.Recordset _RS = null, _RS12 = null;
                       
                        //Utilities.Message("Please wait GRPO creation is in progress....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        _Form.Items.Item("BtnBG").Enabled = false;
                        string oSql = "select T0.\"CardCode\",T0.\"U_BGApp\" " +
                                      " from OCRD T0 where T0.\"CardCode\" = '" + _Form.DataSources.DBDataSources.Item("OCRD").GetValue("CardCode", 0).ToString() + "'";
                        Utilities.ExecuteSQL(ref _RS, oSql);
                        _BPCode = _RS.Fields.Item("CardCode").Value.ToString();
                        string _BPApp = _RS.Fields.Item("U_BGApp").Value.ToString();

                   
                        if (_BPApp.ToString().Trim() == "Yes")
                        {

                            _Form.Items.Item("BtnBG").Enabled = true;
                        }
                        else
                        {
                            _Form.Items.Item("BtnBG").Enabled = false;
                            Utilities.Message("This customer is not eligible for Bank Gurantee Setup", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return;
                        }
                        flag = true;
                    }
                    //switch (pVal.EventType)
                    //{
                    //    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                    //        if (pVal.ItemUID == "txtSbGrp1" && pVal.CharPressed == 9 && Form.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                    //        {
                    //            _Form = Utilities.Application.SBO_Application.Forms.ActiveForm;
                    //            _EditText = (SAPbouiCOM.EditText)_Form.Items.Item("txtSbGrp1").Specific;
                    //            if (_EditText.Value.ToString() == "" && pVal.ItemUID == "txtSbGrp1")
                    //            {
                    //                DataSet ds = LoadCompanySetting();
                    //                string ServerName = ds.Tables[0].Rows[0]["ServerName"].ToString().Trim();
                    //                string UserID = ds.Tables[0].Rows[0]["UserID"].ToString().Trim();
                    //                string Password = ds.Tables[0].Rows[0]["Password"].ToString().Trim();
                    //                string CurrentSchema = ds.Tables[0].Rows[0]["CurrentSchema"].ToString().Trim();
                    //                _conn = new HanaConnection("Server=" + ServerName + ";UserID=" + UserID + ";Password=" + Password + ";Current Schema=" + CurrentSchema + "");
                    //                this._Object = new ClsChooseFromList();
                    //                ((ClsChooseFromList)this._Object).FormTitle = "List of SubGroup1";
                    //                ((ClsChooseFromList)this._Object).OpenCFL(this, pVal, ref oBubbleEvent, "select \"U_SubGrpCodL1\",\"U_SubGrpNameL1\" from \"@ITM_SUBGROUPC1\" where \"U_Status\" = 'Y'", strcon, false);
                    //            }

                    //        }

                    //        if (pVal.ItemUID == "txtSbGrp2" && pVal.CharPressed == 9 && Form.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                    //        {
                    //            _Form = Utilities.Application.SBO_Application.Forms.ActiveForm;
                    //            _EditText = (SAPbouiCOM.EditText)_Form.Items.Item("txtSbGrp2").Specific;
                    //            if (_EditText.Value.ToString() == "" && pVal.ItemUID == "txtSbGrp2")
                    //            {
                    //                DataSet ds = LoadCompanySetting();
                    //                string ServerName = ds.Tables[0].Rows[0]["ServerName"].ToString().Trim();
                    //                string UserID = ds.Tables[0].Rows[0]["UserID"].ToString().Trim();
                    //                string Password = ds.Tables[0].Rows[0]["Password"].ToString().Trim();
                    //                string CurrentSchema = ds.Tables[0].Rows[0]["CurrentSchema"].ToString().Trim();
                    //                _conn = new HanaConnection("Server=" + ServerName + ";UserID=" + UserID + ";Password=" + Password + ";Current Schema=" + CurrentSchema + "");
                    //                this._Object = new ClsChooseFromList();
                    //                ((ClsChooseFromList)this._Object).FormTitle = "List of SubGroup1";
                    //                ((ClsChooseFromList)this._Object).OpenCFL(this, pVal, ref oBubbleEvent, "select \"U_SubGrpCodL2\",\"U_SubGrpNameL2\" from \"@ITM_SUBGROUP2\" where \"U_Status\" = 'Y' and \"U_SubGrpCodL1\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("txtSbGrp1").Specific).Value.ToString() + "'", strcon, false);
                    //            }

                    //        }

                    //        break;
                    //}
                }
                catch (Exception ex)
                {
                    Utilities.ShowErrorMessage(ex.Message);
                }
            }
            #endregion

            #region After Action
            else
            {
                try
                {
                    switch (pVal.EventType)
                    {

                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                            if (pVal.ItemUID == "BtnBG")
                            {
                                _BPCode = Form.DataSources.DBDataSources.Item("OCRD").GetValue("CardCode", 0).ToString().Trim();
                                _BPName = Form.DataSources.DBDataSources.Item("OCRD").GetValue("CardName", 0).ToString();
                                this._Object = new ClsBankGurantee();
                                Utilities.LoadForm(ref this._Object, Constants.Forms.BG_SETUP);
                                ((ClsBankGurantee)this._Object).FormDefault(_BPCode, _BPName);
                            }
                            break;


                    }
                }
                catch (Exception ex)
                {
                    Utilities.ShowErrorMessage(ex.Message);
                }
            }
            #endregion
        }

        #endregion

        #region Form Default
        internal void FormDefault()
        {
            try
            {

                _ExistItem = _Form.Items.Item("2");
                _NewItem = _Form.Items.Add("BtnBG", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                _NewItem.Left = _ExistItem.Left + _ExistItem.Width + 5;
                _NewItem.Width = _ExistItem.Width + 10;
                _NewItem.Height = _ExistItem.Height;
                _NewItem.Top = _ExistItem.Top;
                _Button = (SAPbouiCOM.Button)_NewItem.Specific;
                _Button.Caption = "BG Setup";

            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
        }
        #endregion

        #region LOAD COMPANY SETTINGS
        public static DataSet LoadCompanySetting()
        {
            XmlDataDocument xmlDoc = new XmlDataDocument();
            DataSet dataSet = new DataSet();
            dataSet.ReadXml("AppConfig.xml");
            //dataSet.ReadXml("App.Config");
            return dataSet;

        }
        #endregion

        #region CHOOSE FROM LIST EVENT

        public override void ChooseFromList_Event()
        {
            try
            {
                _Form.Freeze(true);
                if (this._CFLEvent.ItemUID == "txtSbGrp1")
                {
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtSbGrp1").Specific).Value = _CFLEvent.oCFL_SelectedResult.Rows[0][0].ToString();

                }
                if (this._CFLEvent.ItemUID == "txtSbGrp2")
                {
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtSbGrp2").Specific).Value = _CFLEvent.oCFL_SelectedResult.Rows[0][0].ToString();

                }

            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.ToString());
            }
            finally
            {
                _Form.Freeze(false);
            }
        }

        #endregion
    }
}
