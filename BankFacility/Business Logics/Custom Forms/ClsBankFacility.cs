using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;

namespace BankFacility
{
    class ClsBankFacility : Base
    {
        #region Variable declaration

        ChooseFromListEvent _SysCFLEvent;
        Conditions _Cons;
        Condition _Con;
        SAPbobsCOM.Recordset _RS = null, _RSDEL = null;
        private SAPbouiCOM.Matrix _Matrix1, _Matrix2, _Matrix3, _Matrix4, _Matrix11, _Matrix22, _Matrix33, _Matrix44, _Matrix5, _Matrix55, _Matrix6, _Matrix66;
        private SAPbouiCOM.DBDataSource _MasterDBDataSource, _ChildDBDataSource1, _ChildDBDataSource2, _ChildDBDataSource3, _ChildDBDataSource4, _ChildDBDataSource11, _ChildDBDataSource22,_ChildDBDataSource33,_ChildDBDataSource44, _ChildDBDataSource5, _ChildDBDataSource55, _ChildDBDataSource6, _ChildDBDataSource66;
        SAPbouiCOM.ComboBox _ComboBox;
        public string _CardCode = string.Empty;
        public string _SODocEntry = string.Empty;
        private int RowToDelete;
        public static double _TotalLoanInterest =0, _TRLoanInterest = 0,_DMLoanInterest = 0,_STLoanInterest = 0,_ODLoanInterest = 0;
        private string _ParentFormUID = string.Empty;
        private int _ParentRow = -1;

        public bool IsViewOnly { get; set; } = false;

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR

        public ClsBankFacility()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsBankFacility()
        {
        }

        #endregion

        #region ITEM EVENT
        public override void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {

            #region BEFORE ACTION ACTION

            if (pVal.Before_Action == true)
            {
                switch (pVal.EventType)
                {

                    #region ITEM PRESS EVENT

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "20")
                        {
                            _Form.PaneLevel = 1;
                        }
                        else if (pVal.ItemUID == "21")
                        {
                            _Form.PaneLevel = 2;
                        }
                        else if (pVal.ItemUID == "22")
                        {
                            _Form.PaneLevel = 3;
                        }
                        else if (pVal.ItemUID == "23")
                        {
                            _Form.PaneLevel = 4;
                        }
                        else if (pVal.ItemUID == "24")
                        {
                            _Form.PaneLevel = 5;
                        }
                        if (pVal.ItemUID == "1" && (_Form.Mode == BoFormMode.fm_ADD_MODE || _Form.Mode == BoFormMode.fm_UPDATE_MODE))
                        {
                            double TotalFacility = Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("5").Specific).Value);
                            double SumOfFacility = Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("55").Specific).Value) + Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("61").Specific).Value) + Convert.ToDouble(((EditText)_Form.Items.Item("67").Specific).Value) + Convert.ToDouble(((EditText)_Form.Items.Item("73").Specific).Value) + Convert.ToDouble(((EditText)_Form.Items.Item("Item_3").Specific).Value) + +Convert.ToDouble(((EditText)_Form.Items.Item("Item_9").Specific).Value);
                            if (TotalFacility < SumOfFacility)
                            {
                                Utilities.Message("Sum of Facilities amount can not greater than Total Facility Amount", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }
                            //string facilityCode = _MasterDBDataSource.GetValue("U_BnkFactNo", 0).ToString();
                            //string sql = $@"UPDATE ""@BANKFACC1"" SET ""U_Status"" = 'Active' WHERE ""U_Facility"" = '{facilityCode}' ";
                            //if (TotalFacility > 0)
                            //{
                            //    string facilityCode = _MasterDBDataSource.GetValue("U_BnkFactNo", 0).Trim();
                            //    string docEntry = _MasterDBDataSource.GetValue("DocEntry", 0).Trim();

                            //    if (!string.IsNullOrEmpty(facilityCode))
                            //    {
                            //        // You must use a Recordset to run SQL in SAP B1
                            //        SAPbobsCOM.Recordset rsUpdate = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            //        try
                            //        {
                            //            string sql = $@"UPDATE ""@BANKFACC1"" SET ""U_Status"" = 'Active' WHERE ""U_Facility"" = '{facilityCode.Replace("'", "''")}' AND ""U_DocEntry"" = '{docEntry}' ";
                            //            rsUpdate.DoQuery(sql);
                            //        }
                            //        catch (Exception ex)
                            //        {
                            //            Utilities.ShowErrorMessage("SQL Update Error: " + ex.Message);
                            //        }
                            //        finally
                            //        {
                            //            System.Runtime.InteropServices.Marshal.ReleaseComObject(rsUpdate);
                            //        }
                            //    }
                            //}

                        }

                        break;

                    #endregion

                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        if ((pVal.ItemUID == "55" || pVal.ItemUID == "61" || pVal.ItemUID == "67" || pVal.ItemUID == "73") && pVal.CharPressed == 9)
                        {
                            double TotalFacility = Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("5").Specific).Value);
                            double SumOfFacility = Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("55").Specific).Value) + Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("61").Specific).Value) + Convert.ToDouble(((EditText)_Form.Items.Item("67").Specific).Value) + Convert.ToDouble(((EditText)_Form.Items.Item("73").Specific).Value);
                            if(TotalFacility < SumOfFacility)
                            {
                                Utilities.Message("Sum of Facilities amount can not greater than Total Facility Amount", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }

                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                        Func<string, double> GetVal = (alias) => {
                            double v = 0;
                            double.TryParse(_MasterDBDataSource.GetValue(alias, 0).Trim(), out v);
                            return v;
                        };
                        

                        // 3. Update the UI Level (DataSource)
                        //_MasterDBDataSource.SetValue("U_RemFact", 0, remTot.ToString(System.Globalization.CultureInfo.InvariantCulture));
                       
                        if (pVal.ItemUID == "5")
                        {
                            double remTot = GetVal("U_TotFactAmt") - GetVal("U_TotFctAvld");
                            _MasterDBDataSource.SetValue("U_RemFact", 0, remTot.ToString(System.Globalization.CultureInfo.InvariantCulture));
                        }
                        if (pVal.ItemUID == "55")
                        {
                            double remLC = GetVal("U_LCFactAmt") - GetVal("U_LCFctAvld");
                            _MasterDBDataSource.SetValue("U_LCRemFact", 0, remLC.ToString(System.Globalization.CultureInfo.InvariantCulture));
                        }
                        if (pVal.ItemUID == "61")
                        {
                            double remSTL = GetVal("U_STLFactAmt") - GetVal("U_STLFctAvld");
                            _MasterDBDataSource.SetValue("U_STLRemFact", 0, remSTL.ToString(System.Globalization.CultureInfo.InvariantCulture));
                        }
                        if (pVal.ItemUID == "67")
                        {
                            double remOD = GetVal("U_ODFactAmt") - GetVal("U_ODFctAvld");
                            _MasterDBDataSource.SetValue("U_ODRemFact", 0, remOD.ToString(System.Globalization.CultureInfo.InvariantCulture));
                        }
                        if (pVal.ItemUID == "73")
                        {
                            double remDM = GetVal("U_DMFactAmt") - GetVal("U_DMFctAvld");
                            _MasterDBDataSource.SetValue("U_DMRemFact", 0, remDM.ToString(System.Globalization.CultureInfo.InvariantCulture));
                        }
                        if (pVal.ItemUID == "Item_3")
                        {
                            double remBG = GetVal("U_BGFactAmt") - GetVal("U_BGFctAvld");
                            _MasterDBDataSource.SetValue("U_BGRemFact", 0, remBG.ToString(System.Globalization.CultureInfo.InvariantCulture));
                        }
                        if (pVal.ItemUID == "Item_9")
                        {
                            double remTRM = GetVal("U_TRMFactAmt") - GetVal("U_TRMFctAvld");
                            _MasterDBDataSource.SetValue("U_TRMRemFact", 0, remTRM.ToString(System.Globalization.CultureInfo.InvariantCulture));
                        }
                        break;


                    #region et_CHOOSE_FROM_LIST

                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        if (pVal.ItemUID == "10")
                        {
                            this.FilteGLAccount("");
                        }
                        if (pVal.ItemUID == "11")
                        {
                            this.FilteGLAccount("");
                        }
                        if (pVal.ItemUID == "18")
                        {
                            this.FilteGLAccount("");
                        }
                        if (pVal.ItemUID == "19")
                        {
                            this.FilteGLAccount("");
                        }
                        if (pVal.ItemUID == "78")
                        {
                            this.FilteGLAccount("");
                        }
                        if (pVal.ItemUID == "79")
                        {
                            this.FilteGLAccount("");
                        }
                        if (pVal.ItemUID == "80")
                        {
                            this.FilteGLAccount("");
                        }
                        if (pVal.ItemUID == "81")
                        {
                            this.FilteGLAccount("");
                        }
                        break;

                        #endregion

                }
            }

            #endregion

            #region After ACTION ACTION

            else if (pVal.Before_Action == false)
            {
                switch (pVal.EventType)
                {
                    #region et_ITEM_PRESSED

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

                        if (_Form.Mode != BoFormMode.fm_ADD_MODE)
                        {
                            switch (pVal.ItemUID)
                            {
                                case "21": // TR Loan Details Tab
                                    _Form.Freeze(true);
                                    TRLoanLoadData();
                                    TRLoanInterestLoadData();
                                    _Form.Freeze(false);
                                    break;

                                case "22": // Demand Loan Details Tab
                                    _Form.Freeze(true);
                                    DMLoanLoadData();
                                    DMLoanInterestLoadData();
                                    _Form.Freeze(false);
                                    break;

                                case "23": // Short Term Loan Details Tab
                                    _Form.Freeze(true);
                                    STLoanLoadData();
                                    STLoanInterestLoadData();
                                    _Form.Freeze(false);
                                    break;

                                case "24": // Over Draft Details Tab
                                    _Form.Freeze(true);
                                    ODLoadData();
                                    ODInterestLoadData();
                                    _Form.Freeze(false);
                                    break;

                                case "Item_31": // Bridge Gap & Term Loan Details Tab
                                    _Form.Freeze(true);
                                    BridgeGapLoadData();
                                    BridgeGapInterestLoadData();
                                    
                                    _Form.Freeze(false);
                                    break;
                                case "Item_32":
                                    _Form.Freeze(true);
                                    TermLoanLoadData();
                                    TermLoanInterestLoadData();
                                    _Form.Freeze(false);
                                    break;
                            }
                        }

                        if (pVal.ItemUID == "97")
                        {
                            Utilities.Message("Please wait operation is in progress....", SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                            TRLoanLoadData();
                            TRLoanInterestLoadData();
                            DMLoanLoadData();
                            DMLoanInterestLoadData();
                            STLoanLoadData();
                            STLoanInterestLoadData();
                            ODLoadData();
                            ODInterestLoadData();

                            _TotalLoanInterest = _TRLoanInterest + _DMLoanInterest + _STLoanInterest + _ODLoanInterest;
                            ((SAPbouiCOM.EditText)_Form.Items.Item("8").Specific).Value = _TotalLoanInterest.ToString();

                            if (_Form.Mode != BoFormMode.fm_ADD_MODE && pVal.InnerEvent == false)
                            {
                                Utilities.Application.SBO_Application.ActivateMenuItem("1304");
                                _Form.Items.Item("97").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                _Form.Mode = BoFormMode.fm_UPDATE_MODE;
                            }
                            ((SAPbouiCOM.EditText)_Form.Items.Item("6").Specific).Value = (Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("56").Specific).Value) + Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("62").Specific).Value) + Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("68").Specific).Value) + Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("74").Specific).Value) ).ToString();
                            ((SAPbouiCOM.EditText)_Form.Items.Item("7").Specific).Value = (Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("5").Specific).Value) - Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("6").Specific).Value)).ToString();
                            Utilities.Message("Operation completed sucessfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        }
                        if (pVal.ItemUID == "1" && _Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            _Form.Items.Item("12").Enabled = false;

                        }
                        break;

                    #endregion

                    
                    #region Set CFL Values
                    case BoEventTypes.et_CHOOSE_FROM_LIST:
                        _SysCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                        if (pVal.ItemUID == "10" && _SysCFLEvent.SelectedObjects != null)
                        {
                            _MasterDBDataSource.SetValue("U_TRLoanAct", 0, _SysCFLEvent.SelectedObjects.GetValue("AcctCode", 0).ToString().Trim());
                            _MasterDBDataSource.SetValue("U_TRLoanActN", 0, _SysCFLEvent.SelectedObjects.GetValue("AcctName", 0).ToString().Trim());

                        }
                        if (pVal.ItemUID == "11" && _SysCFLEvent.SelectedObjects != null)
                        {
                            _MasterDBDataSource.SetValue("U_DMLoanAct", 0, _SysCFLEvent.SelectedObjects.GetValue("AcctCode", 0).ToString().Trim());
                            _MasterDBDataSource.SetValue("U_DMLoanActN", 0, _SysCFLEvent.SelectedObjects.GetValue("AcctName", 0).ToString().Trim());
                        }
                        if (pVal.ItemUID == "18" && _SysCFLEvent.SelectedObjects != null)
                        {
                            _MasterDBDataSource.SetValue("U_STLAcct", 0, _SysCFLEvent.SelectedObjects.GetValue("AcctCode", 0).ToString().Trim());
                            _MasterDBDataSource.SetValue("U_STLAcctN", 0, _SysCFLEvent.SelectedObjects.GetValue("AcctName", 0).ToString().Trim());
                        }
                        if (pVal.ItemUID == "19" && _SysCFLEvent.SelectedObjects != null)
                        {
                            _MasterDBDataSource.SetValue("U_ODAcct", 0, _SysCFLEvent.SelectedObjects.GetValue("AcctCode", 0).ToString().Trim());
                            _MasterDBDataSource.SetValue("U_ODAcctN", 0, _SysCFLEvent.SelectedObjects.GetValue("AcctName", 0).ToString().Trim());
                        }
                         if (pVal.ItemUID == "78" && _SysCFLEvent.SelectedObjects != null)
                        {
                            _MasterDBDataSource.SetValue("U_ITRLoanAct", 0, _SysCFLEvent.SelectedObjects.GetValue("AcctCode", 0).ToString().Trim());
                            _MasterDBDataSource.SetValue("U_ITRLoanActN", 0, _SysCFLEvent.SelectedObjects.GetValue("AcctName", 0).ToString().Trim());

                        }
                        if (pVal.ItemUID == "80" && _SysCFLEvent.SelectedObjects != null)
                        {
                            _MasterDBDataSource.SetValue("U_ISTLAcct", 0, _SysCFLEvent.SelectedObjects.GetValue("AcctCode", 0).ToString().Trim());
                            _MasterDBDataSource.SetValue("U_ISTLAcctN", 0, _SysCFLEvent.SelectedObjects.GetValue("AcctName", 0).ToString().Trim());
                        }
                        if (pVal.ItemUID == "79" && _SysCFLEvent.SelectedObjects != null)
                        {
                            _MasterDBDataSource.SetValue("U_IDMLoanAct", 0, _SysCFLEvent.SelectedObjects.GetValue("AcctCode", 0).ToString().Trim());
                            _MasterDBDataSource.SetValue("U_IDMLoanActN", 0, _SysCFLEvent.SelectedObjects.GetValue("AcctName", 0).ToString().Trim());
                           
                        }
                        if (pVal.ItemUID == "81" && _SysCFLEvent.SelectedObjects != null)
                        {
                            _MasterDBDataSource.SetValue("U_IODAcct", 0, _SysCFLEvent.SelectedObjects.GetValue("AcctCode", 0).ToString().Trim());
                            _MasterDBDataSource.SetValue("U_IODAcctN", 0, _SysCFLEvent.SelectedObjects.GetValue("AcctName", 0).ToString().Trim());
                        }
                        if (pVal.ItemUID == "3" && _SysCFLEvent.SelectedObjects != null)
                        {
                            _MasterDBDataSource.SetValue("U_BnkCode", 0, _SysCFLEvent.SelectedObjects.GetValue("BankCode", 0).ToString().Trim());
                            _MasterDBDataSource.SetValue("U_BnkName", 0, _SysCFLEvent.SelectedObjects.GetValue("BankName", 0).ToString().Trim());

                        }


                        break;
                        #endregion

                }
            }

            #endregion

        }

        #endregion

        #region MENU EVENT
        public override void Menu_Event(ref SAPbouiCOM.MenuEvent pVal, ref bool oBubbleEvent)
        {
            #region BEFORE ACTION
            if (pVal.BeforeAction)
            {

            }
            #endregion

            #region AFTER ACTION
            else
            {
                try
                {
                    Form.Freeze(true);
                    switch (pVal.MenuUID)
                    {
                        case Constants.System_Menus.mnu_FIND:
                            {
                                _Form.Items.Item("12").Enabled = true;
                            }
                            break;

                        case Constants.System_Menus.mnu_FIRST:
                            BrowseData();
                            break;

                        case Constants.System_Menus.mnu_LAST:
                            BrowseData();
                            break;

                        case Constants.System_Menus.mnu_NEXT:
                            BrowseData();
                            break;

                        case Constants.System_Menus.mnu_PREVIOUS:
                            BrowseData();
                            break;

                    }
                }
                catch (Exception ex)
                {
                    Utilities.ShowErrorMessage(ex.Message);
                }
                finally { Form.Freeze(false); }
            }
            #endregion
        }
        #endregion    

        #region FORM DATA EVENT
        public override void FormData_Event(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool oBubbleEvent)
        {
            #region AFTER ACTION
            if (!BusinessObjectInfo.BeforeAction)
            {
                switch (BusinessObjectInfo.EventType)
                {
                    #region DATA LOAD
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                        try
                        {
                            _Form.Freeze(true);
                            BrowseData();

                        }
                        catch (Exception ex)
                        {
                            Utilities.ShowErrorMessage(ex.Message);
                        }
                        finally
                        {
                            _Form.Freeze(false);
                        }
                        break;

                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                        // This triggers when the user clicks 'Update' and it succeeds
                        if (BusinessObjectInfo.ActionSuccess && !string.IsNullOrEmpty(_ParentFormUID))
                        {
                            // Get current values from the form to send back to the parent
                            string docEntry = _MasterDBDataSource.GetValue("DocEntry", 0).Trim();
                            string total = _MasterDBDataSource.GetValue("U_TotFactAmt", 0).Trim();
                            string availed = _MasterDBDataSource.GetValue("U_TotFctAvld", 0).Trim();
                            string balance = _MasterDBDataSource.GetValue("U_RemFact", 0).Trim();

                            // Update the parent row
                            UpdateParentFormDocEntry(_ParentFormUID, _ParentRow, docEntry, total, availed, balance);
                        }
                        break;
                        #endregion
                }
            }
            #endregion

            #region BEFORE ACTION
            else
            {
                switch (BusinessObjectInfo.EventType)
                {
                    #region DATA LOAD
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:

                        break;
                        #endregion


                }
            }
            #endregion
        }
        #endregion

        #region FILTER BP Type
        internal void FilteGLAccount(string CFLId)
        {

            _Form.Freeze(true);
            try
            {
                _Cons = new Conditions();


                _Con = _Cons.Add();
                _Con.BracketOpenNum = 1;
                _Con.Alias = "Postable";
                _Con.Operation = BoConditionOperation.co_EQUAL;
                _Con.CondVal = "Y";
                _Con.BracketCloseNum = 1;

                _Form.ChooseFromLists.Item("GLAcct1").SetConditions(_Cons);
                _Form.ChooseFromLists.Item("GLAcct2").SetConditions(_Cons);
                _Form.ChooseFromLists.Item("GLAcct3").SetConditions(_Cons);
                _Form.ChooseFromLists.Item("GLAcct4").SetConditions(_Cons);
                _Form.ChooseFromLists.Item("GLAcct5").SetConditions(_Cons);

            }


            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
            }

        }
        #endregion

        #region From Default
        internal void FormDefault()
        {
            try
            {
                _Form.Freeze(true);
                _Form.DataBrowser.BrowseBy = "12";
                _Matrix1 = (SAPbouiCOM.Matrix)Form.Items.Item("42").Specific;
                _Matrix2 = (SAPbouiCOM.Matrix)Form.Items.Item("43").Specific;
                _Matrix3 = (SAPbouiCOM.Matrix)Form.Items.Item("44").Specific;
                _Matrix4 = (SAPbouiCOM.Matrix)Form.Items.Item("1000001").Specific;
                _Matrix5 = (SAPbouiCOM.Matrix)Form.Items.Item("Item_33").Specific;
                _Matrix6 = (SAPbouiCOM.Matrix)Form.Items.Item("Item_35").Specific;

                _Matrix11 = (SAPbouiCOM.Matrix)Form.Items.Item("84").Specific;
                _Matrix22 = (SAPbouiCOM.Matrix)Form.Items.Item("85").Specific;
                _Matrix33 = (SAPbouiCOM.Matrix)Form.Items.Item("86").Specific;
                _Matrix44 = (SAPbouiCOM.Matrix)Form.Items.Item("87").Specific;
                _Matrix55 = (SAPbouiCOM.Matrix)Form.Items.Item("Item_34").Specific;
                _Matrix66 = (SAPbouiCOM.Matrix)Form.Items.Item("Item_36").Specific;


                _MasterDBDataSource = _Form.DataSources.DBDataSources.Item("@BANKFACTH");
                _ChildDBDataSource1 = _Form.DataSources.DBDataSources.Item("@BANKFACTC1");
                _ChildDBDataSource2 = _Form.DataSources.DBDataSources.Item("@BANKFACTC2");
                _ChildDBDataSource3 = _Form.DataSources.DBDataSources.Item("@BANKFACTC3");
                _ChildDBDataSource4 = _Form.DataSources.DBDataSources.Item("@BANKFACTC4");
                _ChildDBDataSource11 = _Form.DataSources.DBDataSources.Item("@BANKFACTC11");
                _ChildDBDataSource22 = _Form.DataSources.DBDataSources.Item("@BANKFACTC22");
                _ChildDBDataSource33 = _Form.DataSources.DBDataSources.Item("@BANKFACTC33");
                _ChildDBDataSource44 = _Form.DataSources.DBDataSources.Item("@BANKFACTC44");
                _ChildDBDataSource5 = _Form.DataSources.DBDataSources.Item("@BANKFACTC5");
                _ChildDBDataSource6 = _Form.DataSources.DBDataSources.Item("@BANKFACTC6");
                _ChildDBDataSource55 = _Form.DataSources.DBDataSources.Item("@BANKFACTC55");
                _ChildDBDataSource66 = _Form.DataSources.DBDataSources.Item("@BANKFACTC66");


                ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value = Utilities.getMaxColumnValueNum("@BANKFACTH", "U_DocNum");
                //_Matrix1.AddRow(1, -1);
                //_Matrix2.AddRow(1, -1);
                //_Matrix3.AddRow(1, -1);
                //_Matrix4.AddRow(1, -1);
                //_Matrix11.AddRow(1, -1);
                //_Matrix22.AddRow(1, -1);
                //_Matrix33.AddRow(1, -1);
                //_Matrix44.AddRow(1, -1);

                _Matrix1.Columns.Item("V_7").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix1.Columns.Item("V_4").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix1.Columns.Item("V_12").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix1.Columns.Item("V_7").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix1.Columns.Item("V_11").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix1.Columns.Item("V_10").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

                _Matrix11.Columns.Item("V_2").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix11.Columns.Item("V_1").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

                _Matrix2.Columns.Item("V_5").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix2.Columns.Item("V_2").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix2.Columns.Item("V_0").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix2.Columns.Item("V_13").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix2.Columns.Item("V_12").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

                _Matrix22.Columns.Item("V_2").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix22.Columns.Item("V_1").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

                _Matrix3.Columns.Item("V_5").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix3.Columns.Item("V_2").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix3.Columns.Item("V_0").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix3.Columns.Item("V_1").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix3.Columns.Item("V_12").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

                _Matrix33.Columns.Item("V_4").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix33.Columns.Item("V_3").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

                _Matrix4.Columns.Item("V_5").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix4.Columns.Item("V_2").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix4.Columns.Item("V_0").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix4.Columns.Item("V_13").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix4.Columns.Item("V_12").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

                _Matrix44.Columns.Item("V_2").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix44.Columns.Item("V_1").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

                _Matrix5.Columns.Item("V_5").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix5.Columns.Item("V_2").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix5.Columns.Item("V_0").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix5.Columns.Item("V_13").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix5.Columns.Item("V_12").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

                _Matrix55.Columns.Item("V_2").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix55.Columns.Item("V_1").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

                _Matrix6.Columns.Item("V_5").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix6.Columns.Item("V_2").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix6.Columns.Item("V_0").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix6.Columns.Item("V_13").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix6.Columns.Item("V_12").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

                _Matrix66.Columns.Item("V_2").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix66.Columns.Item("V_1").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

                //_ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("1000002").Specific;
                //_ComboBox.ValidValues.Add("Open", "Open");
                //_ComboBox.ValidValues.Add("Closed", "Closed");
                //_ComboBox.ValidValues.Add("Cancelled", "Cancelled");
                //_ComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("1000002").Specific;

                // Add this block to prevent the duplicate value error
                if (_ComboBox.ValidValues.Count == 0)
                {
                    _ComboBox.ValidValues.Add("Open", "Open");
                    _ComboBox.ValidValues.Add("Closed", "Closed");
                    _ComboBox.ValidValues.Add("Cancelled", "Cancelled");
                }

                if (IsViewOnly)
                {

                    ((SAPbouiCOM.EditText)_Form.Items.Item("51").Specific).Value = Utilities.Application.SBO_Application.Company.UserName;
                    _Form.Items.Item("49").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    _Form.Items.Item("12").Enabled = false;
                    _Form.Items.Item("4").Enabled = false;
                    _Form.Items.Item("17").Enabled = false;
                    _Form.Items.Item("51").Enabled = false;
                    _Form.Items.Item("8").Enabled = false;
                    _Form.Items.Item("6").Enabled = false;
                    _Form.Items.Item("7").Enabled = false;
                    _Form.Items.Item("56").Enabled = false;
                    _Form.Items.Item("57").Enabled = false;
                    _Form.Items.Item("62").Enabled = false;
                    _Form.Items.Item("63").Enabled = false;
                    _Form.Items.Item("68").Enabled = false;
                    _Form.Items.Item("69").Enabled = false;
                    _Form.Items.Item("74").Enabled = false;
                    _Form.Items.Item("75").Enabled = false;
                    _Form.Items.Item("89").Enabled = false;
                    _Form.Items.Item("90").Enabled = false;
                    _Form.Items.Item("91").Enabled = false;
                    _Form.Items.Item("92").Enabled = false;
                    _Form.Items.Item("93").Enabled = false;
                    _Form.Items.Item("94").Enabled = false;
                    _Form.Items.Item("95").Enabled = false;
                    _Form.Items.Item("96").Enabled = false;


                    _Form.Items.Item("Item_27").Enabled = false;
                    _Form.Items.Item("Item_28").Enabled = false;
                    _Form.Items.Item("Item_29").Enabled = false;
                    _Form.Items.Item("Item_30").Enabled = false;

                    _Form.Items.Item("10").Enabled = false;
                    _Form.Items.Item("18").Enabled = false;
                    _Form.Items.Item("11").Enabled = false;
                    _Form.Items.Item("19").Enabled = false;
                    _Form.Items.Item("Item_21").Enabled = false;
                    _Form.Items.Item("Item_22").Enabled = false;
                    _Form.Items.Item("78").Enabled = false;
                    _Form.Items.Item("80").Enabled = false;
                    _Form.Items.Item("79").Enabled = false;
                    _Form.Items.Item("81").Enabled = false;
                    _Form.Items.Item("Item_25").Enabled = false;
                    _Form.Items.Item("Item_26").Enabled = false;

                    _Form.Items.Item("txtDocE").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                    _Form.Items.Item("49").Enabled = false;
                    _Form.Items.Item("50").Enabled = false;
                    _Form.Items.Item("1000002").Enabled = false;
                    _Form.Items.Item("5").Enabled = false;
                    _Form.Items.Item("55").Enabled = false;
                    _Form.Items.Item("61").Enabled = false;
                    _Form.Items.Item("67").Enabled = false;
                    _Form.Items.Item("73").Enabled = false;
                    _Form.Items.Item("Item_3").Enabled = false;
                    _Form.Items.Item("Item_9").Enabled = false;
                    _Form.Items.Item("82").Enabled = false;


                    _Form.Items.Item("20").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else
                {

                    ((SAPbouiCOM.EditText)_Form.Items.Item("51").Specific).Value = Utilities.Application.SBO_Application.Company.UserName;
                    _Form.Items.Item("49").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    _Form.Items.Item("12").Enabled = false;
                    _Form.Items.Item("4").Enabled = false;
                    _Form.Items.Item("17").Enabled = false;
                    _Form.Items.Item("51").Enabled = false;
                    _Form.Items.Item("8").Enabled = false;
                    _Form.Items.Item("6").Enabled = false;
                    _Form.Items.Item("7").Enabled = false;
                    _Form.Items.Item("56").Enabled = false;
                    _Form.Items.Item("57").Enabled = false;
                    _Form.Items.Item("62").Enabled = false;
                    _Form.Items.Item("63").Enabled = false;
                    _Form.Items.Item("68").Enabled = false;
                    _Form.Items.Item("69").Enabled = false;
                    _Form.Items.Item("74").Enabled = false;
                    _Form.Items.Item("75").Enabled = false;
                    _Form.Items.Item("89").Enabled = false;
                    _Form.Items.Item("90").Enabled = false;
                    _Form.Items.Item("91").Enabled = false;
                    _Form.Items.Item("92").Enabled = false;
                    _Form.Items.Item("93").Enabled = false;
                    _Form.Items.Item("94").Enabled = false;
                    _Form.Items.Item("95").Enabled = false;
                    _Form.Items.Item("96").Enabled = false;


                    _Form.Items.Item("Item_27").Enabled = false;
                    _Form.Items.Item("Item_28").Enabled = false;
                    _Form.Items.Item("Item_29").Enabled = false;
                    _Form.Items.Item("Item_30").Enabled = false;

                    _Form.Items.Item("10").Enabled = false;
                    _Form.Items.Item("18").Enabled = false;
                    _Form.Items.Item("11").Enabled = false;
                    _Form.Items.Item("19").Enabled = false;
                    _Form.Items.Item("Item_21").Enabled = false;
                    _Form.Items.Item("Item_22").Enabled = false;
                    _Form.Items.Item("78").Enabled = false;
                    _Form.Items.Item("80").Enabled = false;
                    _Form.Items.Item("79").Enabled = false;
                    _Form.Items.Item("81").Enabled = false;
                    _Form.Items.Item("Item_25").Enabled = false;
                    _Form.Items.Item("Item_26").Enabled = false;


                    _Form.Items.Item("20").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
                _Form.Visible = true;
            }
        }


        #endregion

        #region Loan Data 

        #region TRLoanLoadData
        public void TRLoanLoadData()
        {
            string facility = _MasterDBDataSource.GetValue("U_BnkFactNo", 0).Trim();
            _Matrix1.Clear();
            _ChildDBDataSource1.Clear();
            string _StrQuery = "select T0.\"Number\" as \"JENo\",T0.\"TaxDate\" as \"PostingDate\",T2.\"PrcCode\",T2.\"PrcName\",T4.\"AcctCode\", " +
                                           " T4.\"AcctName\",T2.\"ValidFrom\",T2.\"ValidTo\",T1.\"Debit\",T1.\"Credit\", " +
                                           "SUM(T1.\"Credit\") as \"Availed Facility\", T5.\"U_LoanAmount\" " +
                                           " from OJDT T0 inner join JDT1 T1 ON T0.\"TransId\" = T1.\"TransId\" " +
                                           " inner join OPRC T2 ON T2.\"PrcCode\" = T1.\"OcrCode5\" " +
                                           " inner join \"@BANKFACTH\" T3 ON  T3.\"U_TRLoanAct\" = T1.\"Account\" " +
                                           " inner join OACT T4 ON T4.\"AcctCode\" = T1.\"Account\"  " +
                                           " inner join \"@LOANMASTERC\"  T5 ON T3.\"U_BnkFactNo\" = T5.\"U_BankFacNo\" " +
                                           " where T2.\"GrpCode\" = 'TR' " +
                                            " and T3.\"U_BnkFactNo\" = '" + facility.Replace("'", "''") + "' " +
                                           "group by T0.\"Number\",T0.\"TaxDate\",T2.\"PrcCode\",T2.\"PrcName\",T4.\"AcctCode\", " +
                                           " T4.\"AcctName\",T2.\"ValidFrom\",T2.\"ValidTo\",T1.\"Debit\",T1.\"Credit\", T5.\"U_LoanAmount\" " +
                                           "ORDER BY T2.\"PrcCode\",T1.\"Debit\" ASC ";
            Utilities.ExecuteSQL(ref _RS, _StrQuery);
            //if (_RS.RecordCount > 0)
            //{
                ((SAPbouiCOM.EditText)_Form.Items.Item("56").Specific).Value = _RS.Fields.Item("Availed Facility").Value.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("57").Specific).Value = (Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("55").Specific).Value) - Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("56").Specific).Value)).ToString();
            string _StrQueryDel = "delete from \"@BANKFACTC1\" where \"DocEntry\" in (select \"DocEntry\" from \"@BANKFACTH\" where \"U_DocNum\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value.ToString() + "')";
            Utilities.ExecuteSQL(ref _RSDEL, _StrQueryDel);

            for (int i = 0; i < _RS.RecordCount; i++)
                {
                    _Matrix1.AddRow(1, -1);
                    _Matrix1.FlushToDataSource();
                    _ChildDBDataSource1.SetValue("U_SrNo", i, (i + 1).ToString());
                    _ChildDBDataSource1.SetValue("U_JENo", i, _RS.Fields.Item("JENo").Value.ToString());
                    _ChildDBDataSource1.SetValue("U_JEDt", i, DateTime.Parse(_RS.Fields.Item("PostingDate").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource1.SetValue("U_CCNo", i, _RS.Fields.Item("PrcCode").Value.ToString());
                    _ChildDBDataSource1.SetValue("U_CCName", i, _RS.Fields.Item("PrcName").Value.ToString());
                    _ChildDBDataSource1.SetValue("U_GLCode", i, _RS.Fields.Item("AcctCode").Value.ToString());
                    _ChildDBDataSource1.SetValue("U_GLName", i, _RS.Fields.Item("AcctName").Value.ToString());
                    _ChildDBDataSource1.SetValue("U_LCAmt", i, _RS.Fields.Item("U_LoanAmount").Value.ToString());
                    _ChildDBDataSource1.SetValue("U_LCStDt", i, DateTime.Parse(_RS.Fields.Item("ValidFrom").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource1.SetValue("U_LCEnDt", i, DateTime.Parse(_RS.Fields.Item("ValidTo").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource1.SetValue("U_DrAmt", i, _RS.Fields.Item("Debit").Value.ToString());
                    _ChildDBDataSource1.SetValue("U_CrAmt", i, _RS.Fields.Item("Credit").Value.ToString());

                    _Matrix1.LoadFromDataSource();
                    _RS.MoveNext();
                }
               
            //}
        }
        #endregion

        #region TRLoanInterestLoadData
        public void TRLoanInterestLoadData()
        {
            string facility = _MasterDBDataSource.GetValue("U_BnkFactNo", 0).Trim();

            _Matrix11.Clear();
            _ChildDBDataSource11.Clear();
            string _StrQuery = "select T0.\"Number\" as \"JENo\",T0.\"TaxDate\" as \"PostingDate\",T2.\"PrcCode\",T2.\"PrcName\",T4.\"AcctCode\", " +
                                           " T4.\"AcctName\",T2.\"ValidFrom\",T2.\"ValidTo\",T1.\"Debit\",T1.\"Credit\", " +
                                           "SUM(T1.\"Debit\") as \"Availed Facility\" " +
                                           " from OJDT T0 inner join JDT1 T1 ON T0.\"TransId\" = T1.\"TransId\" " +
                                           " inner join OPRC T2 ON T2.\"PrcCode\" = T1.\"OcrCode5\" " +
                                           " inner join \"@BANKFACTH\" T3 ON T3.\"U_ITRLoanAct\" = T1.\"Account\" " +
                                           " inner join \"@LOANMASTERC\"  T5 ON T3.\"U_BnkFactNo\" = T5.\"U_BankFacNo\" " +
                                           " inner join OACT T4 ON T4.\"AcctCode\" = T1.\"Account\" " +
                                            " where T2.\"GrpCode\" = 'TR' " +
                                            " and T3.\"U_BnkFactNo\" = '" + facility.Replace("'", "''") + "' " +
                                           " group by T0.\"Number\",T0.\"TaxDate\",T2.\"PrcCode\",T2.\"PrcName\",T4.\"AcctCode\", " +
                                           " T4.\"AcctName\",T2.\"ValidFrom\",T2.\"ValidTo\",T1.\"Debit\",T1.\"Credit\" " +
                                           "ORDER BY T2.\"PrcCode\",T1.\"Debit\" ASC ";
            Utilities.ExecuteSQL(ref _RS, _StrQuery);
            //if (_RS.RecordCount > 0)
            //{
                _TRLoanInterest = Convert.ToDouble(_RS.Fields.Item("Availed Facility").Value.ToString());

            string _StrQueryDel = "delete from \"@BANKFACTC11\" where \"DocEntry\" in (select \"DocEntry\" from \"@BANKFACTH\" where \"U_DocNum\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value.ToString() + "')";
            Utilities.ExecuteSQL(ref _RSDEL, _StrQueryDel);
            for (int i = 0; i < _RS.RecordCount; i++)
                {
                    _Matrix11.AddRow(1, -1);
                    _Matrix11.FlushToDataSource();
                    _ChildDBDataSource11.SetValue("U_SrNo", i, (i + 1).ToString());
                    _ChildDBDataSource11.SetValue("LineId", i, (i + 1).ToString());
                    _ChildDBDataSource11.SetValue("U_JENo", i, _RS.Fields.Item("JENo").Value.ToString());
                    _ChildDBDataSource11.SetValue("U_JEDt", i, DateTime.Parse(_RS.Fields.Item("PostingDate").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource11.SetValue("U_CCNo", i, _RS.Fields.Item("PrcCode").Value.ToString());
                    _ChildDBDataSource11.SetValue("U_CCName", i, _RS.Fields.Item("PrcName").Value.ToString());
                    _ChildDBDataSource11.SetValue("U_GLCode", i, _RS.Fields.Item("AcctCode").Value.ToString());
                    _ChildDBDataSource11.SetValue("U_GLName", i, _RS.Fields.Item("AcctName").Value.ToString());
                    _ChildDBDataSource11.SetValue("U_DrAmt", i, _RS.Fields.Item("Debit").Value.ToString());
                    _ChildDBDataSource11.SetValue("U_CrAmt", i, _RS.Fields.Item("Credit").Value.ToString());
                    // _ChildDBDataSource2.SetValue("U_BookedBy", i, _RS.Fields.Item("Credit").Value.ToString());

                    _Matrix11.LoadFromDataSource();
                    _RS.MoveNext();
                }
            //}
        }

        #endregion

        #region DMLoanLoadData
        public void DMLoanLoadData()
        {
            string facility = _MasterDBDataSource.GetValue("U_BnkFactNo", 0).Trim();
            _Matrix2.Clear();
            _ChildDBDataSource2.Clear();

            //TR Loan
            string _StrQuery = "select T0.\"Number\" as \"JENo\",T0.\"TaxDate\" as \"PostingDate\",T2.\"PrcCode\",T2.\"PrcName\",T4.\"AcctCode\", " +
                                           " T4.\"AcctName\",T5.\"U_LoanAmount\",T2.\"ValidFrom\",T2.\"ValidTo\",T1.\"Debit\",T1.\"Credit\", " +
                                           "SUM(T1.\"Credit\") as \"Availed Facility\" " +
                                           " from OJDT T0 inner join JDT1 T1 ON T0.\"TransId\" = T1.\"TransId\" " +
                                           " inner join OPRC T2 ON T2.\"PrcCode\" = T1.\"OcrCode5\" " +
                                           " inner join \"@BANKFACTH\" T3 ON T3.\"U_DMLoanAct\" = T1.\"Account\" " +
                                           " inner join \"@LOANMASTERC\"  T5 ON T3.\"U_BnkFactNo\" = T5.\"U_BankFacNo\" " +
                                           " inner join OACT T4 ON T4.\"AcctCode\" = T1.\"Account\" " +
                                            " where T2.\"GrpCode\" = 'DM' " +
                                            " and T3.\"U_BnkFactNo\" = '" + facility.Replace("'", "''") + "' " +
                                           "group by T0.\"Number\",T0.\"TaxDate\",T2.\"PrcCode\",T2.\"PrcName\",T4.\"AcctCode\", " +
                                          " T4.\"AcctName\",T5.\"U_LoanAmount\",T2.\"ValidFrom\",T2.\"ValidTo\",T1.\"Debit\",T1.\"Credit\" " +
                                          "ORDER BY T2.\"PrcCode\",T1.\"Debit\" ASC ";
            Utilities.ExecuteSQL(ref _RS, _StrQuery);
            //if (_RS.RecordCount > 0)
            //{
                ((SAPbouiCOM.EditText)_Form.Items.Item("74").Specific).Value = _RS.Fields.Item("Availed Facility").Value.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("75").Specific).Value = (Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("73").Specific).Value) - Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("74").Specific).Value)).ToString();


            string _StrQueryDel = "delete from \"@BANKFACTC2\" where \"DocEntry\" in (select \"DocEntry\" from \"@BANKFACTH\" where \"U_DocNum\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value.ToString() + "')";
            Utilities.ExecuteSQL(ref _RSDEL, _StrQueryDel);
            for (int i = 0; i < _RS.RecordCount; i++)
                {
                    _Matrix2.AddRow(1, -1);
                    _Matrix2.FlushToDataSource();
                    _ChildDBDataSource2.SetValue("U_SrNo", i, (i + 1).ToString());
                    _ChildDBDataSource2.SetValue("U_JENo", i, _RS.Fields.Item("JENo").Value.ToString());
                    _ChildDBDataSource2.SetValue("U_JEDt", i, DateTime.Parse(_RS.Fields.Item("PostingDate").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource2.SetValue("U_CCNo", i, _RS.Fields.Item("PrcCode").Value.ToString());
                    _ChildDBDataSource2.SetValue("U_CCName", i, _RS.Fields.Item("PrcName").Value.ToString());
                    _ChildDBDataSource2.SetValue("U_GLCode", i, _RS.Fields.Item("AcctCode").Value.ToString());
                    _ChildDBDataSource2.SetValue("U_GLName", i, _RS.Fields.Item("AcctName").Value.ToString());
                    _ChildDBDataSource2.SetValue("U_DMAmt", i, _RS.Fields.Item("U_LoanAmount").Value.ToString());
                    _ChildDBDataSource2.SetValue("U_DMStDt", i, DateTime.Parse(_RS.Fields.Item("ValidFrom").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource2.SetValue("U_DMEnDt", i, DateTime.Parse(_RS.Fields.Item("ValidTo").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource2.SetValue("U_DrAmt", i, _RS.Fields.Item("Debit").Value.ToString());
                    _ChildDBDataSource2.SetValue("U_CrAmt", i, _RS.Fields.Item("Credit").Value.ToString());

                    _Matrix2.LoadFromDataSource();
                    _RS.MoveNext();
                }
            //}
        }

        #endregion

        #region DMLoanInterestLoadData
        public void DMLoanInterestLoadData()
        {
            string facility = _MasterDBDataSource.GetValue("U_BnkFactNo", 0).Trim();
            _Matrix22.Clear();
            _ChildDBDataSource22.Clear();
            //TR Loan
            string _StrQuery = "select T0.\"Number\" as \"JENo\",T0.\"TaxDate\" as \"PostingDate\",T2.\"PrcCode\",T2.\"PrcName\",T4.\"AcctCode\", " +
                                           " T4.\"AcctName\",T2.\"ValidFrom\",T2.\"ValidTo\",T1.\"Debit\",T1.\"Credit\", " +
                                           "SUM(T1.\"Debit\") as \"Availed Facility\" " +
                                           " from OJDT T0 inner join JDT1 T1 ON T0.\"TransId\" = T1.\"TransId\" " +
                                           " inner join OPRC T2 ON T2.\"PrcCode\" = T1.\"OcrCode5\" " +
                                           " inner join \"@BANKFACTH\" T3 ON T3.\"U_IDMLoanAct\" = T1.\"Account\" " +
                                            " inner join \"@LOANMASTERC\"  T5 ON T3.\"U_BnkFactNo\" = T5.\"U_BankFacNo\" " +
                                           " inner join OACT T4 ON T4.\"AcctCode\" = T1.\"Account\" " +
                                           " where T2.\"GrpCode\" = 'DM' " +
                                            " and T3.\"U_BnkFactNo\" = '" + facility.Replace("'", "''") + "' " +
                                           " group by T0.\"Number\",T0.\"TaxDate\",T2.\"PrcCode\",T2.\"PrcName\",T4.\"AcctCode\", " +
                                           " T4.\"AcctName\",T2.\"ValidFrom\",T2.\"ValidTo\",T1.\"Debit\",T1.\"Credit\" " +
                                           "ORDER BY T2.\"PrcCode\",T1.\"Debit\" ASC ";
            Utilities.ExecuteSQL(ref _RS, _StrQuery);
            //if (_RS.RecordCount > 0)
            //{
                _DMLoanInterest = Convert.ToDouble(_RS.Fields.Item("Availed Facility").Value.ToString());

            string _StrQueryDel = "delete from \"@BANKFACTC22\" where \"DocEntry\" in (select \"DocEntry\" from \"@BANKFACTH\" where \"U_DocNum\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value.ToString() + "')";
            Utilities.ExecuteSQL(ref _RSDEL, _StrQueryDel);
            for (int i = 0; i < _RS.RecordCount; i++)
                {
                    _Matrix22.AddRow(1, -1);
                    _Matrix22.FlushToDataSource();
                    _ChildDBDataSource22.SetValue("U_SrNo", i, (i + 1).ToString());
                    _ChildDBDataSource22.SetValue("U_JENo", i, _RS.Fields.Item("JENo").Value.ToString());
                    _ChildDBDataSource22.SetValue("U_JEDt", i, DateTime.Parse(_RS.Fields.Item("PostingDate").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource22.SetValue("U_CCNo", i, _RS.Fields.Item("PrcCode").Value.ToString());
                    _ChildDBDataSource22.SetValue("U_CCName", i, _RS.Fields.Item("PrcName").Value.ToString());
                    _ChildDBDataSource22.SetValue("U_GLCode", i, _RS.Fields.Item("AcctCode").Value.ToString());
                    _ChildDBDataSource22.SetValue("U_GLName", i, _RS.Fields.Item("AcctName").Value.ToString());
                    _ChildDBDataSource22.SetValue("U_DrAmt", i, _RS.Fields.Item("Debit").Value.ToString());
                    _ChildDBDataSource22.SetValue("U_CrAmt", i, _RS.Fields.Item("Credit").Value.ToString());
                    // _ChildDBDataSource2.SetValue("U_BookedBy", i, _RS.Fields.Item("Credit").Value.ToString());

                    _Matrix22.LoadFromDataSource();
                    _RS.MoveNext();
                }
            //}
        }
        #endregion

        #region STLoanLoadData
        public void STLoanLoadData()
        {
            string facility = _MasterDBDataSource.GetValue("U_BnkFactNo", 0).Trim();
            _Matrix3.Clear();
            _ChildDBDataSource3.Clear();
            //TR Loan
            string _StrQuery = "select T0.\"Number\" as \"JENo\",T0.\"TaxDate\" as \"PostingDate\",T2.\"PrcCode\",T2.\"PrcName\",T4.\"AcctCode\", " +
                                           " T4.\"AcctName\",T5.\"U_LoanAmount\",T2.\"ValidFrom\",T2.\"ValidTo\",T1.\"Debit\",T1.\"Credit\", " +
                                            "SUM(T1.\"Credit\") as \"Availed Facility\" " +
                                           " from OJDT T0 inner join JDT1 T1 ON T0.\"TransId\" = T1.\"TransId\" " +
                                           " inner join OPRC T2 ON T2.\"PrcCode\" = T1.\"OcrCode4\" " +
                                           " inner join \"@BANKFACTH\" T3 ON T3.\"U_STLAcct\" = T1.\"Account\" " +
                                            " inner join \"@LOANMASTERC\"  T5 ON T3.\"U_BnkFactNo\" = T5.\"U_BankFacNo\" " +
                                           " inner join OACT T4 ON T4.\"AcctCode\" = T1.\"Account\" " +
                                            " where T2.\"GrpCode\" = 'STR' " +
                                            " and T3.\"U_BnkFactNo\" = '" + facility.Replace("'", "''") + "' " +
                                           " group by T0.\"Number\",T0.\"TaxDate\",T2.\"PrcCode\",T2.\"PrcName\",T4.\"AcctCode\", " +
                                           " T4.\"AcctName\",T5.\"U_LoanAmount\",T2.\"ValidFrom\",T2.\"ValidTo\",T1.\"Debit\",T1.\"Credit\" " +
                                           "ORDER BY T2.\"PrcCode\",T1.\"Debit\" ASC ";
            Utilities.ExecuteSQL(ref _RS, _StrQuery);
            //if (_RS.RecordCount > 0)
            //{
                ((SAPbouiCOM.EditText)_Form.Items.Item("62").Specific).Value = _RS.Fields.Item("Availed Facility").Value.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("63").Specific).Value = (Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("61").Specific).Value) - Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("62").Specific).Value)).ToString();

            string _StrQueryDel = "delete from \"@BANKFACTC3\" where \"DocEntry\" in (select \"DocEntry\" from \"@BANKFACTH\" where \"U_DocNum\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value.ToString() + "')";
            Utilities.ExecuteSQL(ref _RSDEL, _StrQueryDel);
            for (int i = 0; i < _RS.RecordCount; i++)
                {
                    _Matrix3.AddRow(1, -1);
                    _Matrix3.FlushToDataSource();
                    _ChildDBDataSource3.SetValue("U_SrNo", i, (i + 1).ToString());
                    _ChildDBDataSource3.SetValue("U_JENo", i, _RS.Fields.Item("JENo").Value.ToString());
                    _ChildDBDataSource3.SetValue("U_JEDt", i, DateTime.Parse(_RS.Fields.Item("PostingDate").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource3.SetValue("U_CCNo", i, _RS.Fields.Item("PrcCode").Value.ToString());
                    _ChildDBDataSource3.SetValue("U_CCName", i, _RS.Fields.Item("PrcName").Value.ToString());
                    _ChildDBDataSource3.SetValue("U_GLCode", i, _RS.Fields.Item("AcctCode").Value.ToString());
                    _ChildDBDataSource3.SetValue("U_GLName", i, _RS.Fields.Item("AcctName").Value.ToString());
                    _ChildDBDataSource3.SetValue("U_STLAmt", i, _RS.Fields.Item("U_LoanAmount").Value.ToString());
                    _ChildDBDataSource3.SetValue("U_STLStDt", i, DateTime.Parse(_RS.Fields.Item("ValidFrom").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource3.SetValue("U_STLEnDt", i, DateTime.Parse(_RS.Fields.Item("ValidTo").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource3.SetValue("U_DrAmt", i, _RS.Fields.Item("Debit").Value.ToString());
                    _ChildDBDataSource3.SetValue("U_CrAmt", i, _RS.Fields.Item("Credit").Value.ToString());

                    _Matrix3.LoadFromDataSource();
                    _RS.MoveNext();
                }
            //}
        }
        #endregion

        #region STLoanInterestLoadData
        public void STLoanInterestLoadData()
        {
            string facility = _MasterDBDataSource.GetValue("U_BnkFactNo", 0).Trim();
            _Matrix33.Clear();
            _ChildDBDataSource33.Clear();
            //TR Loan
            string _StrQuery = "select T0.\"Number\" as \"JENo\",T0.\"TaxDate\" as \"PostingDate\",T2.\"PrcCode\",T2.\"PrcName\",T4.\"AcctCode\", " +
                                           " T4.\"AcctName\",T5.\"U_LoanAmount\",T2.\"ValidFrom\",T2.\"ValidTo\",T1.\"Debit\",T1.\"Credit\", " +
                                           " SUM(T1.\"Debit\") as \"Availed Facility\" " +
                                           " from OJDT T0 inner join JDT1 T1 ON T0.\"TransId\" = T1.\"TransId\" " +
                                           " inner join OPRC T2 ON T2.\"PrcCode\" = T1.\"OcrCode4\" " +
                                           " inner join \"@BANKFACTH\" T3 ON T3.\"U_ISTLAcct\" = T1.\"Account\" " +
                                            " inner join \"@LOANMASTERC\"  T5 ON T3.\"U_BnkFactNo\" = T5.\"U_BankFacNo\" " +
                                           " inner join OACT T4 ON T4.\"AcctCode\" = T1.\"Account\" " +
                                            " where T2.\"GrpCode\" = 'STR' " +
                                            " and T3.\"U_BnkFactNo\" = '" + facility.Replace("'", "''") + "' " +
                                          " group by T0.\"Number\",T0.\"TaxDate\",T2.\"PrcCode\",T2.\"PrcName\",T4.\"AcctCode\", " +
                                           " T4.\"AcctName\",T5.\"U_LoanAmount\",T2.\"ValidFrom\",T2.\"ValidTo\",T1.\"Debit\",T1.\"Credit\" " +
                                           "ORDER BY T2.\"PrcCode\",T1.\"Debit\" ASC ";
            Utilities.ExecuteSQL(ref _RS, _StrQuery);
            //if (_RS.RecordCount > 0)
            //{
                _STLoanInterest = Convert.ToDouble(_RS.Fields.Item("Availed Facility").Value.ToString());

            string _StrQueryDel = "delete from \"@BANKFACTC33\" where \"DocEntry\" in (select \"DocEntry\" from \"@BANKFACTH\" where \"U_DocNum\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value.ToString() + "')";
            Utilities.ExecuteSQL(ref _RSDEL, _StrQueryDel);
            for (int i = 0; i < _RS.RecordCount; i++)
                {
                    _Matrix33.AddRow(1, -1);
                    _Matrix33.FlushToDataSource();
                    _ChildDBDataSource33.SetValue("U_SrNo", i, (i + 1).ToString());
                    _ChildDBDataSource33.SetValue("U_JENo", i, _RS.Fields.Item("JENo").Value.ToString());
                    _ChildDBDataSource33.SetValue("U_JEDt", i, DateTime.Parse(_RS.Fields.Item("PostingDate").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource33.SetValue("U_CCNo", i, _RS.Fields.Item("PrcCode").Value.ToString());
                    _ChildDBDataSource33.SetValue("U_CCName", i, _RS.Fields.Item("PrcName").Value.ToString());
                    _ChildDBDataSource33.SetValue("U_GLCode", i, _RS.Fields.Item("AcctCode").Value.ToString());
                    _ChildDBDataSource33.SetValue("U_GLName", i, _RS.Fields.Item("AcctName").Value.ToString());
                    _ChildDBDataSource33.SetValue("U_DrAmt", i, _RS.Fields.Item("Debit").Value.ToString());
                    _ChildDBDataSource33.SetValue("U_CrAmt", i, _RS.Fields.Item("Credit").Value.ToString());
                    // _ChildDBDataSource2.SetValue("U_BookedBy", i, _RS.Fields.Item("Credit").Value.ToString());

                    _Matrix33.LoadFromDataSource();
                    _RS.MoveNext();
                }
            //}
        }
        #endregion

        #region ODLoadData
        public void ODLoadData()
        {
            string facility = _MasterDBDataSource.GetValue("U_BnkFactNo", 0).Trim();
            _Matrix4.Clear();
            _ChildDBDataSource4.Clear();
            //TR Loan
            string _StrQuery = "select T0.\"Number\" as \"JENo\",T0.\"TaxDate\" as \"PostingDate\",T2.\"PrcCode\",T2.\"PrcName\",T4.\"AcctCode\", " +
                                           " T4.\"AcctName\",T5.\"U_LoanAmount\",T2.\"ValidFrom\",T2.\"ValidTo\",T1.\"Debit\",T1.\"Credit\", " +
                                           "SUM(T1.\"Credit\") as \"Availed Facility\" " +
                                           " from OJDT T0 inner join JDT1 T1 ON T0.\"TransId\" = T1.\"TransId\" " +
                                           " inner join OPRC T2 ON T2.\"PrcCode\" = T1.\"OcrCode4\" " +
                                           " inner join \"@BANKFACTH\" T3 ON T3.\"U_ODAcct\" = T1.\"Account\" " +
                                           " inner join \"@LOANMASTERC\"  T5 ON T3.\"U_BnkFactNo\" = T5.\"U_BankFacNo\" " +
                                           " inner join OACT T4 ON T4.\"AcctCode\" = T1.\"Account\" " +
                                           " where T2.\"GrpCode\" = 'OD' " +
                                            " and T3.\"U_BnkFactNo\" = '" + facility.Replace("'", "''") + "' " +
                                           " group by T0.\"Number\",T0.\"TaxDate\",T2.\"PrcCode\",T2.\"PrcName\",T4.\"AcctCode\", " +
                                           " T4.\"AcctName\",T5.\"U_LoanAmount\",T2.\"ValidFrom\",T2.\"ValidTo\",T1.\"Debit\",T1.\"Credit\" " +
                                           "ORDER BY T2.\"PrcCode\",T1.\"Debit\" ASC ";
            Utilities.ExecuteSQL(ref _RS, _StrQuery);
            //if (_RS.RecordCount > 0)
            //{
                ((SAPbouiCOM.EditText)_Form.Items.Item("68").Specific).Value = _RS.Fields.Item("Availed Facility").Value.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("69").Specific).Value = (Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("67").Specific).Value) - Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("68").Specific).Value)).ToString();

            string _StrQueryDel = "delete from \"@BANKFACTC4\" where \"DocEntry\" in (select \"DocEntry\" from \"@BANKFACTH\" where \"U_DocNum\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value.ToString() + "')";
            Utilities.ExecuteSQL(ref _RSDEL, _StrQueryDel);
            for (int i = 0; i < _RS.RecordCount; i++)
                {
                    _Matrix4.AddRow(1, -1);
                    _Matrix4.FlushToDataSource();
                    _ChildDBDataSource4.SetValue("U_SrNo", i, (i + 1).ToString());
                    _ChildDBDataSource4.SetValue("U_JENo", i, _RS.Fields.Item("JENo").Value.ToString());
                    _ChildDBDataSource4.SetValue("U_JEDt", i, DateTime.Parse(_RS.Fields.Item("PostingDate").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource4.SetValue("U_CCNo", i, _RS.Fields.Item("PrcCode").Value.ToString());
                    _ChildDBDataSource4.SetValue("U_CCName", i, _RS.Fields.Item("PrcName").Value.ToString());
                    _ChildDBDataSource4.SetValue("U_GLCode", i, _RS.Fields.Item("AcctCode").Value.ToString());
                    _ChildDBDataSource4.SetValue("U_GLName", i, _RS.Fields.Item("AcctName").Value.ToString());
                    _ChildDBDataSource4.SetValue("U_ODAmt", i, _RS.Fields.Item("U_LoanAmount").Value.ToString());
                    _ChildDBDataSource4.SetValue("U_ODStDt", i, DateTime.Parse(_RS.Fields.Item("ValidFrom").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource4.SetValue("U_ODEnDt", i, DateTime.Parse(_RS.Fields.Item("ValidTo").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource4.SetValue("U_DrAmt", i, _RS.Fields.Item("Debit").Value.ToString());
                    _ChildDBDataSource4.SetValue("U_CrAmt", i, _RS.Fields.Item("Credit").Value.ToString());

                    _Matrix4.LoadFromDataSource();
                    _RS.MoveNext();
                }
            //}
        }
        #endregion

        #region ODInterestLoadData
        public void ODInterestLoadData()
        {
            string facility = _MasterDBDataSource.GetValue("U_BnkFactNo", 0).Trim();
            _Matrix44.Clear();
            _ChildDBDataSource44.Clear();
            //TR Loan
            string _StrQuery = "select T0.\"Number\" as \"JENo\",T0.\"TaxDate\" as \"PostingDate\",T2.\"PrcCode\",T2.\"PrcName\",T4.\"AcctCode\", " +
                                           " T4.\"AcctName\",T5.\"U_LoanAmount\",T2.\"ValidFrom\",T2.\"ValidTo\",T1.\"Debit\",T1.\"Credit\", " +
                                            " SUM(T1.\"Debit\") as \"Availed Facility\" " +
                                           " from OJDT T0 inner join JDT1 T1 ON T0.\"TransId\" = T1.\"TransId\" " +
                                           " inner join OPRC T2 ON T2.\"PrcCode\" = T1.\"OcrCode4\" " +
                                           " inner join \"@BANKFACTH\" T3 ON T3.\"U_IODAcct\" = T1.\"Account\" " +
                                            " inner join \"@LOANMASTERC\"  T5 ON T3.\"U_BnkFactNo\" = T5.\"U_BankFacNo\" " +
                                           " inner join OACT T4 ON T4.\"AcctCode\" = T1.\"Account\"  " +
                                           " where T2.\"GrpCode\" = 'OD' " +
                                            " and T3.\"U_BnkFactNo\" = '" + facility.Replace("'", "''") + "' " +
                                           " group by T0.\"Number\",T0.\"TaxDate\",T2.\"PrcCode\",T2.\"PrcName\",T4.\"AcctCode\", " +
                                           " T4.\"AcctName\",T5.\"U_LoanAmount\",T2.\"ValidFrom\",T2.\"ValidTo\",T1.\"Debit\",T1.\"Credit\" " +
                                           "ORDER BY T2.\"PrcCode\",T1.\"Debit\" ASC ";
            Utilities.ExecuteSQL(ref _RS, _StrQuery);
            //if (_RS.RecordCount > 0)
            //{
                _ODLoanInterest = Convert.ToDouble(_RS.Fields.Item("Availed Facility").Value.ToString());

            string _StrQueryDel = "delete from \"@BANKFACTC44\" where \"DocEntry\" in (select \"DocEntry\" from \"@BANKFACTH\" where \"U_DocNum\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value.ToString() + "')";
            Utilities.ExecuteSQL(ref _RSDEL, _StrQueryDel);
            for (int i = 0; i < _RS.RecordCount; i++)
                {
                    _Matrix44.AddRow(1, -1);
                    _Matrix44.FlushToDataSource();
                    _ChildDBDataSource44.SetValue("U_SrNo", i, (i + 1).ToString());
                    _ChildDBDataSource44.SetValue("U_JENo", i, _RS.Fields.Item("JENo").Value.ToString());
                    _ChildDBDataSource44.SetValue("U_JEDt", i, DateTime.Parse(_RS.Fields.Item("PostingDate").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource44.SetValue("U_CCNo", i, _RS.Fields.Item("PrcCode").Value.ToString());
                    _ChildDBDataSource44.SetValue("U_CCName", i, _RS.Fields.Item("PrcName").Value.ToString());
                    _ChildDBDataSource44.SetValue("U_GLCode", i, _RS.Fields.Item("AcctCode").Value.ToString());
                    _ChildDBDataSource44.SetValue("U_GLName", i, _RS.Fields.Item("AcctName").Value.ToString());
                    _ChildDBDataSource44.SetValue("U_DrAmt", i, _RS.Fields.Item("Debit").Value.ToString());
                    _ChildDBDataSource44.SetValue("U_CrAmt", i, _RS.Fields.Item("Credit").Value.ToString());
                    // _ChildDBDataSource2.SetValue("U_BookedBy", i, _RS.Fields.Item("Credit").Value.ToString());

                    _Matrix44.LoadFromDataSource();
                    _RS.MoveNext();
                }
            //}
        }
        #endregion

        #region BridgeGapLoadData
        public void BridgeGapLoadData()
        {
            string facility = _MasterDBDataSource.GetValue("U_BnkFactNo", 0).Trim();
            _Matrix5.Clear();
            _ChildDBDataSource5.Clear();
            //TR Loan
            string _StrQuery = "select T0.\"Number\" as \"JENo\",T0.\"TaxDate\" as \"PostingDate\",T2.\"PrcCode\",T2.\"PrcName\",T4.\"AcctCode\", " +
                                           " T4.\"AcctName\",T5.\"U_LoanAmount\",T2.\"ValidFrom\",T2.\"ValidTo\",T1.\"Debit\",T1.\"Credit\", " +
                                           "SUM(T1.\"Credit\") as \"Availed Facility\" " +
                                           " from OJDT T0 inner join JDT1 T1 ON T0.\"TransId\" = T1.\"TransId\" " +
                                           " inner join OPRC T2 ON T2.\"PrcCode\" = T1.\"OcrCode4\" " +
                                           " inner join \"@BANKFACTH\" T3 ON T3.\"U_ODAcct\" = T1.\"Account\" " +
                                            " inner join \"@LOANMASTERC\"  T5 ON T3.\"U_BnkFactNo\" = T5.\"U_BankFacNo\" " +
                                           " inner join OACT T4 ON T4.\"AcctCode\" = T1.\"Account\" " +
                                            " where T2.\"GrpCode\" = 'BGL' " +
                                            " and T3.\"U_BnkFactNo\" = '" + facility.Replace("'", "''") + "' " +
                                           " group by T0.\"Number\",T0.\"TaxDate\",T2.\"PrcCode\",T2.\"PrcName\",T4.\"AcctCode\", " +
                                           " T4.\"AcctName\",T5.\"U_LoanAmount\",T2.\"ValidFrom\",T2.\"ValidTo\",T1.\"Debit\",T1.\"Credit\" " +
                                           "ORDER BY T2.\"PrcCode\",T1.\"Debit\" ASC ";
            Utilities.ExecuteSQL(ref _RS, _StrQuery);
            //if (_RS.RecordCount > 0)
            //{
            ((SAPbouiCOM.EditText)_Form.Items.Item("68").Specific).Value = _RS.Fields.Item("Availed Facility").Value.ToString();
            ((SAPbouiCOM.EditText)_Form.Items.Item("69").Specific).Value = (Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("67").Specific).Value) - Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("68").Specific).Value)).ToString();

            string _StrQueryDel = "delete from \"@BANKFACTC5\" where \"DocEntry\" in (select \"DocEntry\" from \"@BANKFACTH\" where \"U_DocNum\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value.ToString() + "')";
            Utilities.ExecuteSQL(ref _RSDEL, _StrQueryDel);
            for (int i = 0; i < _RS.RecordCount; i++)
            {
                _Matrix5.AddRow(1, -1);
                _Matrix5.FlushToDataSource();
                _ChildDBDataSource5.SetValue("U_SrNo", i, (i + 1).ToString());
                _ChildDBDataSource5.SetValue("U_JENo", i, _RS.Fields.Item("JENo").Value.ToString());
                _ChildDBDataSource5.SetValue("U_JEDt", i, DateTime.Parse(_RS.Fields.Item("PostingDate").Value.ToString()).ToString("yyyyMMdd"));
                _ChildDBDataSource5.SetValue("U_CCNo", i, _RS.Fields.Item("PrcCode").Value.ToString());
                _ChildDBDataSource5.SetValue("U_CCName", i, _RS.Fields.Item("PrcName").Value.ToString());
                _ChildDBDataSource5.SetValue("U_GLCode", i, _RS.Fields.Item("AcctCode").Value.ToString());
                _ChildDBDataSource5.SetValue("U_GLName", i, _RS.Fields.Item("AcctName").Value.ToString());
                _ChildDBDataSource5.SetValue("U_ODAmt", i, _RS.Fields.Item("U_LoanAmount").Value.ToString());
                _ChildDBDataSource5.SetValue("U_ODStDt", i, DateTime.Parse(_RS.Fields.Item("ValidFrom").Value.ToString()).ToString("yyyyMMdd"));
                _ChildDBDataSource5.SetValue("U_ODEnDt", i, DateTime.Parse(_RS.Fields.Item("ValidTo").Value.ToString()).ToString("yyyyMMdd"));
                _ChildDBDataSource5.SetValue("U_DrAmt", i, _RS.Fields.Item("Debit").Value.ToString());
                _ChildDBDataSource5.SetValue("U_CrAmt", i, _RS.Fields.Item("Credit").Value.ToString());

                _Matrix5.LoadFromDataSource();
                _RS.MoveNext();
            }
            //}
        }
        #endregion

        #region BridgeGapInterestLoadData
        public void BridgeGapInterestLoadData()
        {
            string facility = _MasterDBDataSource.GetValue("U_BnkFactNo", 0).Trim();
            _Matrix55.Clear();
            _ChildDBDataSource55.Clear();
            //TR Loan
            string _StrQuery = "select T0.\"Number\" as \"JENo\",T0.\"TaxDate\" as \"PostingDate\",T2.\"PrcCode\",T2.\"PrcName\",T4.\"AcctCode\", " +
                                           " T4.\"AcctName\",T5.\"U_LoanAmount\",T2.\"ValidFrom\",T2.\"ValidTo\",T1.\"Debit\",T1.\"Credit\", " +
                                            " SUM(T1.\"Debit\") as \"Availed Facility\" " +
                                           " from OJDT T0 inner join JDT1 T1 ON T0.\"TransId\" = T1.\"TransId\" " +
                                           " inner join OPRC T2 ON T2.\"PrcCode\" = T1.\"OcrCode4\" " +
                                           " inner join \"@BANKFACTH\" T3 ON T3.\"U_IODAcct\" = T1.\"Account\" " +
                                           " inner join \"@LOANMASTERC\"  T5 ON T3.\"U_BnkFactNo\" = T5.\"U_BankFacNo\" " +
                                           " inner join OACT T4 ON T4.\"AcctCode\" = T1.\"Account\"  " +
                                            " where T2.\"GrpCode\" = 'BGL' " +
                                            " and T3.\"U_BnkFactNo\" = '" + facility.Replace("'", "''") + "' " +
                                           " group by T0.\"Number\",T0.\"TaxDate\",T2.\"PrcCode\",T2.\"PrcName\",T4.\"AcctCode\", " +
                                           " T4.\"AcctName\",T5.\"U_LoanAmount\",T2.\"ValidFrom\",T2.\"ValidTo\",T1.\"Debit\",T1.\"Credit\" " +
                                           "ORDER BY T2.\"PrcCode\",T1.\"Debit\" ASC ";
            Utilities.ExecuteSQL(ref _RS, _StrQuery);
            //if (_RS.RecordCount > 0)
            //{
            _ODLoanInterest = Convert.ToDouble(_RS.Fields.Item("Availed Facility").Value.ToString());

            string _StrQueryDel = "delete from \"@BANKFACTC55\" where \"DocEntry\" in (select \"DocEntry\" from \"@BANKFACTH\" where \"U_DocNum\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value.ToString() + "')";
            Utilities.ExecuteSQL(ref _RSDEL, _StrQueryDel);
            for (int i = 0; i < _RS.RecordCount; i++)
            {
                _Matrix55.AddRow(1, -1);
                _Matrix55.FlushToDataSource();
                _ChildDBDataSource55.SetValue("U_SrNo", i, (i + 1).ToString());
                _ChildDBDataSource55.SetValue("U_JENo", i, _RS.Fields.Item("JENo").Value.ToString());
                _ChildDBDataSource55.SetValue("U_JEDt", i, DateTime.Parse(_RS.Fields.Item("PostingDate").Value.ToString()).ToString("yyyyMMdd"));
                _ChildDBDataSource55.SetValue("U_CCNo", i, _RS.Fields.Item("PrcCode").Value.ToString());
                _ChildDBDataSource55.SetValue("U_CCName", i, _RS.Fields.Item("PrcName").Value.ToString());
                _ChildDBDataSource55.SetValue("U_GLCode", i, _RS.Fields.Item("AcctCode").Value.ToString());
                _ChildDBDataSource55.SetValue("U_GLName", i, _RS.Fields.Item("AcctName").Value.ToString());
                _ChildDBDataSource55.SetValue("U_DrAmt", i, _RS.Fields.Item("Debit").Value.ToString());
                _ChildDBDataSource55.SetValue("U_CrAmt", i, _RS.Fields.Item("Credit").Value.ToString());
                // _ChildDBDataSource2.SetValue("U_BookedBy", i, _RS.Fields.Item("Credit").Value.ToString());

                _Matrix55.LoadFromDataSource();
                _RS.MoveNext();
            }
            //}
        }
        #endregion

        #region TermLoanLoadData
        public void TermLoanLoadData()
        {
            string facility = _MasterDBDataSource.GetValue("U_BnkFactNo", 0).Trim();
            _Matrix6.Clear();
            _ChildDBDataSource6.Clear();
            //TR Loan
            string _StrQuery = "select T0.\"Number\" as \"JENo\",T0.\"TaxDate\" as \"PostingDate\",T2.\"PrcCode\",T2.\"PrcName\",T4.\"AcctCode\", " +
                                           " T4.\"AcctName\",T5.\"U_LoanAmount\",T2.\"ValidFrom\",T2.\"ValidTo\",T1.\"Debit\",T1.\"Credit\", " +
                                           "SUM(T1.\"Credit\") as \"Availed Facility\" " +
                                           " from OJDT T0 inner join JDT1 T1 ON T0.\"TransId\" = T1.\"TransId\" " +
                                           " inner join OPRC T2 ON T2.\"PrcCode\" = T1.\"OcrCode4\" " +
                                           " inner join \"@BANKFACTH\" T3 ON T3.\"U_ODAcct\" = T1.\"Account\" " +
                                           " inner join \"@LOANMASTERC\"  T5 ON T3.\"U_BnkFactNo\" = T5.\"U_BankFacNo\" " +
                                           " inner join OACT T4 ON T4.\"AcctCode\" = T1.\"Account\" " +
                                            " where T2.\"GrpCode\" = 'TRM' " +
                                            " and T3.\"U_BnkFactNo\" = '" + facility.Replace("'", "''") + "' " +
                                           " group by T0.\"Number\",T0.\"TaxDate\",T2.\"PrcCode\",T2.\"PrcName\",T4.\"AcctCode\", " +
                                           " T4.\"AcctName\",T5.\"U_LoanAmount\",T2.\"ValidFrom\",T2.\"ValidTo\",T1.\"Debit\",T1.\"Credit\" " +
                                           "ORDER BY T2.\"PrcCode\",T1.\"Debit\" ASC ";
            Utilities.ExecuteSQL(ref _RS, _StrQuery);
            //if (_RS.RecordCount > 0)
            //{
            ((SAPbouiCOM.EditText)_Form.Items.Item("68").Specific).Value = _RS.Fields.Item("Availed Facility").Value.ToString();
            ((SAPbouiCOM.EditText)_Form.Items.Item("69").Specific).Value = (Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("67").Specific).Value) - Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("68").Specific).Value)).ToString();

            string _StrQueryDel = "delete from \"@BANKFACTC6\" where \"DocEntry\" in (select \"DocEntry\" from \"@BANKFACTH\" where \"U_DocNum\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value.ToString() + "')";
            Utilities.ExecuteSQL(ref _RSDEL, _StrQueryDel);
            for (int i = 0; i < _RS.RecordCount; i++)
            {
                _Matrix6.AddRow(1, -1);
                _Matrix6.FlushToDataSource();
                _ChildDBDataSource6.SetValue("U_SrNo", i, (i + 1).ToString());
                _ChildDBDataSource6.SetValue("U_JENo", i, _RS.Fields.Item("JENo").Value.ToString());
                _ChildDBDataSource6.SetValue("U_JEDt", i, DateTime.Parse(_RS.Fields.Item("PostingDate").Value.ToString()).ToString("yyyyMMdd"));
                _ChildDBDataSource6.SetValue("U_CCNo", i, _RS.Fields.Item("PrcCode").Value.ToString());
                _ChildDBDataSource6.SetValue("U_CCName", i, _RS.Fields.Item("PrcName").Value.ToString());
                _ChildDBDataSource6.SetValue("U_GLCode", i, _RS.Fields.Item("AcctCode").Value.ToString());
                _ChildDBDataSource6.SetValue("U_GLName", i, _RS.Fields.Item("AcctName").Value.ToString());
                _ChildDBDataSource6.SetValue("U_ODAmt", i, _RS.Fields.Item("U_LoanAmount").Value.ToString());
                _ChildDBDataSource6.SetValue("U_ODStDt", i, DateTime.Parse(_RS.Fields.Item("ValidFrom").Value.ToString()).ToString("yyyyMMdd"));
                _ChildDBDataSource6.SetValue("U_ODEnDt", i, DateTime.Parse(_RS.Fields.Item("ValidTo").Value.ToString()).ToString("yyyyMMdd"));
                _ChildDBDataSource6.SetValue("U_DrAmt", i, _RS.Fields.Item("Debit").Value.ToString());
                _ChildDBDataSource6.SetValue("U_CrAmt", i, _RS.Fields.Item("Credit").Value.ToString());

                _Matrix6.LoadFromDataSource();
                _RS.MoveNext();
            }
            //}
        }
        #endregion

        #region TermLoanInterestLoadData
        public void TermLoanInterestLoadData()
        {
            string facility = _MasterDBDataSource.GetValue("U_BnkFactNo", 0).Trim();
            _Matrix66.Clear();
            _ChildDBDataSource66.Clear();
            //TR Loan
            string _StrQuery = "select T0.\"Number\" as \"JENo\",T0.\"TaxDate\" as \"PostingDate\",T2.\"PrcCode\",T2.\"PrcName\",T4.\"AcctCode\", " +
                                           " T4.\"AcctName\",T5.\"U_LoanAmount\",T2.\"ValidFrom\",T2.\"ValidTo\",T1.\"Debit\",T1.\"Credit\", " +
                                            " SUM(T1.\"Debit\") as \"Availed Facility\" " +
                                           " from OJDT T0 inner join JDT1 T1 ON T0.\"TransId\" = T1.\"TransId\" " +
                                           " inner join OPRC T2 ON T2.\"PrcCode\" = T1.\"OcrCode4\" " +
                                           " inner join \"@BANKFACTH\" T3 ON T3.\"U_IODAcct\" = T1.\"Account\" " +
                                           " inner join \"@LOANMASTERC\"  T5 ON T3.\"U_BnkFactNo\" = T5.\"U_BankFacNo\" " +
                                           " inner join OACT T4 ON T4.\"AcctCode\" = T1.\"Account\"  " +
                                            " where T2.\"GrpCode\" = 'TRM' " +
                                            " and T3.\"U_BnkFactNo\" = '" + facility.Replace("'", "''") + "' " +
                                           " group by T0.\"Number\",T0.\"TaxDate\",T2.\"PrcCode\",T2.\"PrcName\",T4.\"AcctCode\", " +
                                           " T4.\"AcctName\",T5.\"U_LoanAmount\",T2.\"ValidFrom\",T2.\"ValidTo\",T1.\"Debit\",T1.\"Credit\" " +
                                           "ORDER BY T2.\"PrcCode\",T1.\"Debit\" ASC ";
            Utilities.ExecuteSQL(ref _RS, _StrQuery);
            //if (_RS.RecordCount > 0)
            //{
            _ODLoanInterest = Convert.ToDouble(_RS.Fields.Item("Availed Facility").Value.ToString());

            string _StrQueryDel = "delete from \"@BANKFACTC66\" where \"DocEntry\" in (select \"DocEntry\" from \"@BANKFACTH\" where \"U_DocNum\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value.ToString() + "')";
            Utilities.ExecuteSQL(ref _RSDEL, _StrQueryDel);
            for (int i = 0; i < _RS.RecordCount; i++)
            {
                _Matrix66.AddRow(1, -1);
                _Matrix66.FlushToDataSource();
                _ChildDBDataSource66.SetValue("U_SrNo", i, (i + 1).ToString());
                _ChildDBDataSource66.SetValue("U_JENo", i, _RS.Fields.Item("JENo").Value.ToString());
                _ChildDBDataSource66.SetValue("U_JEDt", i, DateTime.Parse(_RS.Fields.Item("PostingDate").Value.ToString()).ToString("yyyyMMdd"));
                _ChildDBDataSource66.SetValue("U_CCNo", i, _RS.Fields.Item("PrcCode").Value.ToString());
                _ChildDBDataSource66.SetValue("U_CCName", i, _RS.Fields.Item("PrcName").Value.ToString());
                _ChildDBDataSource66.SetValue("U_GLCode", i, _RS.Fields.Item("AcctCode").Value.ToString());
                _ChildDBDataSource66.SetValue("U_GLName", i, _RS.Fields.Item("AcctName").Value.ToString());
                _ChildDBDataSource66.SetValue("U_DrAmt", i, _RS.Fields.Item("Debit").Value.ToString());
                _ChildDBDataSource66.SetValue("U_CrAmt", i, _RS.Fields.Item("Credit").Value.ToString());
                // _ChildDBDataSource2.SetValue("U_BookedBy", i, _RS.Fields.Item("Credit").Value.ToString());

                _Matrix66.LoadFromDataSource();
                _RS.MoveNext();
            }
            //}
        }
        #endregion

        #endregion

        #region On Form Load if Amounts are changed 
        public void BrowseData()
        {
            try
            {
                if (_Form == null) return;
                _Form.Freeze(true);

                FormDefault();

                // 1. Calculate Days
                string endDateValue = ((SAPbouiCOM.EditText)_Form.Items.Item("16").Specific).Value.ToString();
                string daysBetween = "0";
                if (!string.IsNullOrEmpty(endDateValue))
                {
                    string _StrQuery = $@"SELECT DAYS_BETWEEN(CURRENT_UTCDATE, '{endDateValue}') FROM DUMMY";
                    Utilities.ExecuteSQL(ref _RS, _StrQuery);
                    if (!_RS.EoF) daysBetween = _RS.Fields.Item(0).Value.ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("17").Specific).Value = daysBetween;
                }

                // 2. Perform Calculations
                Func<string, double> GetVal = (alias) => {
                    double v = 0;
                    double.TryParse(_MasterDBDataSource.GetValue(alias, 0).Trim(), out v);
                    return v;
                };

                double remTot = GetVal("U_TotFactAmt") - GetVal("U_TotFctAvld");
                double remLC = GetVal("U_LCFactAmt") - GetVal("U_LCFctAvld");
                double remSTL = GetVal("U_STLFactAmt") - GetVal("U_STLFctAvld");
                double remOD = GetVal("U_ODFactAmt") - GetVal("U_ODFctAvld");
                double remDM = GetVal("U_DMFactAmt") - GetVal("U_DMFctAvld");
                double remBG = GetVal("U_BGFactAmt") - GetVal("U_BGFctAvld");
                double remTRM = GetVal("U_TRMFactAmt") - GetVal("U_TRMFctAvld");

                // 3. Update the UI Level (DataSource)
                //_MasterDBDataSource.SetValue("U_RemFact", 0, remTot.ToString(System.Globalization.CultureInfo.InvariantCulture));
                //_MasterDBDataSource.SetValue("U_LCRemFact", 0, remLC.ToString(System.Globalization.CultureInfo.InvariantCulture));
                //_MasterDBDataSource.SetValue("U_STLRemFact", 0, remSTL.ToString(System.Globalization.CultureInfo.InvariantCulture));
                //_MasterDBDataSource.SetValue("U_ODRemFact", 0, remOD.ToString(System.Globalization.CultureInfo.InvariantCulture));
                //_MasterDBDataSource.SetValue("U_DMRemFact", 0, remDM.ToString(System.Globalization.CultureInfo.InvariantCulture));
                //_MasterDBDataSource.SetValue("U_BGRemFact", 0, remBG.ToString(System.Globalization.CultureInfo.InvariantCulture));
                //_MasterDBDataSource.SetValue("U_TRMRemFact", 0, remTRM.ToString(System.Globalization.CultureInfo.InvariantCulture));

                double SumOfFacility = Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("56").Specific).Value) + Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("62").Specific).Value) + Convert.ToDouble(((EditText)_Form.Items.Item("68").Specific).Value) + Convert.ToDouble(((EditText)_Form.Items.Item("74").Specific).Value) + Convert.ToDouble(((EditText)_Form.Items.Item("Item_4").Specific).Value) + +Convert.ToDouble(((EditText)_Form.Items.Item("Item_10").Specific).Value);

                SetValueIfChanged("U_RemFact", remTot);
                SetValueIfChanged("U_LCRemFact", remLC);
                SetValueIfChanged("U_STLRemFact", remSTL);
                SetValueIfChanged("U_ODRemFact", remOD);
                SetValueIfChanged("U_DMRemFact", remDM);
                SetValueIfChanged("U_BGRemFact", remBG);
                SetValueIfChanged("U_TRMRemFact", remTRM);
                SetValueIfChanged("U_TotFctAvld", SumOfFacility);

                // Update the DATABASE Level (SQL)
                // Only update if we have a valid DocEntry (not in Add mode)
                string docEntry = _MasterDBDataSource.GetValue("DocEntry", 0).Trim();
                //if (!string.IsNullOrEmpty(docEntry) && docEntry != "0")
                //{
                //    string sqlUpdate = $@"
                //UPDATE ""@BANKFACTH"" SET 
                //""U_RemFact"" = {remTot}, 
                //""U_LCRemFact"" = {remLC}, 
                //""U_STLRemFact"" = {remSTL}, 
                //""U_ODRemFact"" = {remOD}, 
                //""U_DMRemFact"" = {remDM}, 
                //""U_BGRemFact"" = {remBG}, 
                //""U_TRMRemFact"" = {remTRM},
                //""U_RemDays"" = '{daysBetween}'
                //WHERE ""DocEntry"" = {docEntry}";

                //    SAPbobsCOM.Recordset rsUpd = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                //    rsUpd.DoQuery(sqlUpdate);
                //    System.Runtime.InteropServices.Marshal.ReleaseComObject(rsUpd);
                //}


                if (_Form.Mode == BoFormMode.fm_UPDATE_MODE)
                {
                    _Form.Mode = BoFormMode.fm_OK_MODE;
                }
            }
            catch (Exception ex) { /* Log error */ }
            finally { _Form.Freeze(false); }
        }
        #endregion

        #region Helper function to get if value changed
        private void SetValueIfChanged(string alias, double newValue)
        {
            // 1. Get current value from the DataSource
            string currentStr = _MasterDBDataSource.GetValue(alias, 0).Trim();
            double currentVal = 0;
            double.TryParse(currentStr, out currentVal);

            // 2. Only update if the values are different
            // We use Math.Round to avoid tiny decimal differences (0.000001) triggering an update
            if (Math.Round(currentVal, 2) != Math.Round(newValue, 2))
            {
                _MasterDBDataSource.SetValue(alias, 0, newValue.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }
        }
        #endregion

        private void PersistDerivedFields(string docEntry)
        {
            if (string.IsNullOrEmpty(docEntry) || docEntry == "0") return;

            string sql = $@"
        UPDATE ""@BANKFACTH"" SET
            ""U_RemFact""    = IFNULL(""U_TotFactAmt"",0) - IFNULL(""U_TotFctAvld"",0),
            ""U_LCRemFact""  = IFNULL(""U_LCFactAmt"",0)  - IFNULL(""U_LCFctAvld"",0),
            ""U_STLRemFact"" = IFNULL(""U_STLFactAmt"",0) - IFNULL(""U_STLFctAvld"",0),
            ""U_ODRemFact""  = IFNULL(""U_ODFactAmt"",0)  - IFNULL(""U_ODFctAvld"",0),
            ""U_DMRemFact""  = IFNULL(""U_DMFactAmt"",0)  - IFNULL(""U_DMFctAvld"",0),
            ""U_BGRemFact""  = IFNULL(""U_BGFactAmt"",0)  - IFNULL(""U_BGFctAvld"",0),
            ""U_TRMRemFact"" = IFNULL(""U_TRMFactAmt"",0) - IFNULL(""U_TRMFctAvld"",0),
            ""U_RemDays""    = DAYS_BETWEEN(CURRENT_UTCDATE, ""U_EndDate"")
        WHERE ""DocEntry"" = {docEntry}";

            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)
                Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            try { rs.DoQuery(sql); }
            finally { System.Runtime.InteropServices.Marshal.ReleaseComObject(rs); }
        }


        #region Load the data from bank facility on double click
        //public void LoadData(string bankCode, string bankName, string startDate, string endDate, string bankFacilityNo)
        //{
        //    try
        //    {
        //        _Form.Freeze(true);

        //        SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //        string sql = $"SELECT \"DocEntry\" FROM \"@BANKFACTH\" WHERE \"U_BnkFactNo\" = '{bankFacilityNo.Replace("'", "''")}'";
        //        rs.DoQuery(sql);

        //        //if (!rs.EoF)
        //        //{
        //        //    // RECORD EXISTS
        //        //    _Form.Mode = BoFormMode.fm_FIND_MODE;
        //        //    _Form.Items.Item("txtLoanC").Enabled = true;
        //        //    ((SAPbouiCOM.EditText)_Form.Items.Item("txtLoanC").Specific).Value = loanCode;
        //        //    _Form.Items.Item("1").Click();
        //        //    _Form.Items.Item("txtLoanC").Enabled = false;
        //        //}
        //        //else
        //        //{
        //        // RECORD DOES NOT EXIST
        //            _MasterDBDataSource.SetValue("U_BnkCode", 0, bankCode);
        //            _MasterDBDataSource.SetValue("U_BnkName", 0, bankName);
        //            _MasterDBDataSource.SetValue("U_StartDate", 0, startDate);
        //            _MasterDBDataSource.SetValue("U_EndDate", 0, endDate);
        //            _MasterDBDataSource.SetValue("U_BnkFactNo", 0, bankFacilityNo);

        //        //DateTime startD = Convert.ToDateTime(startDate);
        //        //DateTime endD = Convert.ToDateTime(endDate);

        //        //int remDays = (int)(endD - startD).TotalDays;
        //        //_MasterDBDataSource.SetValue("U_RemDays", 0, remDays.ToString());

        //        int remDays = CalculateDateDiff(startDate, endDate);
        //        _MasterDBDataSource.SetValue("U_RemDays", 0, remDays.ToString());

        //        SAPbobsCOM.Recordset rs1 = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //        //string sql1 = $@"SELECT T0.""U_TrLoan"", T0.""U_StlLoan"", T0.""U_DemandLoan"", T0.""U_ODLoan"", 
        //        //T0.""U_TermLoan"", T0.""U_BridgeLoan"", T0.""U_TRInterestAc"", T0.""U_STLInterestAc"", T0.""U_DemandInterestAc"", 
        //        //T0.""U_ODInterestAc"", T0.""U_TermInterestAc"", T0.""U_BridgeInterestAc"" 
        //        //FROM ""@LOANGLAC"" T0 WHERE T0.""Code"" = '{bankCode.Replace("'", "''")}' ";

        //        string sql1 = $@"
        //    SELECT 
        //        T0.""U_TrLoan"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_TrLoan"") AS ""TrLoanName"",
        //        T0.""U_StlLoan"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_StlLoan"") AS ""StlLoanName"",
        //        T0.""U_DemandLoan"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_DemandLoan"") AS ""DemandLoanName"",
        //        T0.""U_ODLoan"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_ODLoan"") AS ""ODLoanName"",
        //        T0.""U_TermLoan"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_TermLoan"") AS ""TermLoanName"",
        //        T0.""U_BridgeLoan"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_BridgeLoan"") AS ""BridgeLoanName"",
        //        T0.""U_TRInterestAc"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_TRInterestAc"") AS ""TRIntName"",
        //        T0.""U_STLInterestAc"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_STLInterestAc"") AS ""STLIntName"",
        //        T0.""U_DemandInterestAc"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_DemandInterestAc"") AS ""DemandIntName"",
        //        T0.""U_ODInterestAc"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_ODInterestAc"") AS ""ODIntName"",
        //        T0.""U_TermInterestAc"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_TermInterestAc"") AS ""TermIntName"",
        //        T0.""U_BridgeInterestAc"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_BridgeInterestAc"") AS ""BridgeIntName""
        //    FROM ""@LOANGLAC"" T0 
        //    WHERE T0.""Code"" = '{bankCode.Replace("'", "''")}'";

        //        rs1.DoQuery(sql1);

        //        _MasterDBDataSource.SetValue("U_TRLoanAct", 0, rs1.Fields.Item("U_TrLoan").Value.ToString());
        //        _MasterDBDataSource.SetValue("U_STLAcct", 0, rs1.Fields.Item("U_StlLoan").Value.ToString());
        //        _MasterDBDataSource.SetValue("U_DMLoanAct", 0, rs1.Fields.Item("U_DemandLoan").Value.ToString());
        //        _MasterDBDataSource.SetValue("U_ODAcct", 0, rs1.Fields.Item("U_ODLoan").Value.ToString());
        //        _MasterDBDataSource.SetValue("U_BrdgLAcct", 0, rs1.Fields.Item("U_BridgeLoan").Value.ToString());
        //        _MasterDBDataSource.SetValue("U_TrmLAcct", 0, rs1.Fields.Item("U_TermLoan").Value.ToString());

        //        _MasterDBDataSource.SetValue("U_TRLoanActN", 0, rs1.Fields.Item("TrLoanName").Value.ToString());
        //        _MasterDBDataSource.SetValue("U_STLAcctN", 0, rs1.Fields.Item("StlLoanName").Value.ToString());
        //        _MasterDBDataSource.SetValue("U_DMLoanActN", 0, rs1.Fields.Item("DemandLoanName").Value.ToString());
        //        _MasterDBDataSource.SetValue("U_ODAcctN", 0, rs1.Fields.Item("ODLoanName").Value.ToString());
        //        _MasterDBDataSource.SetValue("U_BrdgLAcctN", 0, rs1.Fields.Item("BridgeLoanName").Value.ToString());
        //        _MasterDBDataSource.SetValue("U_TrmLAcctN", 0, rs1.Fields.Item("TermLoanName").Value.ToString());

        //        _MasterDBDataSource.SetValue("U_ITRLoanAct", 0, rs1.Fields.Item("U_TRInterestAc").Value.ToString());
        //        _MasterDBDataSource.SetValue("U_ISTLAcct", 0, rs1.Fields.Item("U_STLInterestAc").Value.ToString());
        //        _MasterDBDataSource.SetValue("U_IDMLoanAct", 0, rs1.Fields.Item("U_DemandInterestAc").Value.ToString());
        //        _MasterDBDataSource.SetValue("U_IODAcct", 0, rs1.Fields.Item("U_ODInterestAc").Value.ToString());
        //        _MasterDBDataSource.SetValue("U_IBGAcct", 0, rs1.Fields.Item("U_BridgeInterestAc").Value.ToString());
        //        _MasterDBDataSource.SetValue("U_ITRMAcct", 0, rs1.Fields.Item("U_TermInterestAc").Value.ToString());

        //        _MasterDBDataSource.SetValue("U_ITRLoanActN", 0, rs1.Fields.Item("TRIntName").Value.ToString());
        //        _MasterDBDataSource.SetValue("U_ISTLAcctN", 0, rs1.Fields.Item("STLIntName").Value.ToString());
        //        _MasterDBDataSource.SetValue("U_IDMLoanActN", 0, rs1.Fields.Item("DemandIntName").Value.ToString());
        //        _MasterDBDataSource.SetValue("U_IODAcctN", 0, rs1.Fields.Item("ODIntName").Value.ToString());
        //        _MasterDBDataSource.SetValue("U_IBGAcctN", 0, rs1.Fields.Item("BridgeIntName").Value.ToString());
        //        _MasterDBDataSource.SetValue("U_ITRMAcctN", 0, rs1.Fields.Item("TermIntName").Value.ToString());



        //        //}
        //    }
        //    catch (Exception ex)
        //    {
        //        Utilities.ShowErrorMessage("Error loading data: " + ex.Message);
        //    }
        //    finally
        //    {
        //        _Form.Freeze(false);
        //    }


        //}

        public void LoadData(string bankCode, string bankName, string startDate, string endDate, string bankFacilityNo, string existingDocEntry, string parentUID, int parentRowIdx)
        {
            SAPbobsCOM.Recordset rs = null;
            _ParentFormUID = parentUID;
            _ParentRow = parentRowIdx;
            try
            {
                _Form.Freeze(true);

                rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                bool recordNeedsCreation = string.IsNullOrEmpty(existingDocEntry) || existingDocEntry == "0";

                if (!recordNeedsCreation)
                {
                    // CASE 1: DocEntry already known. Find it immediately.

                    SyncGLAccountsFromMapping(bankCode, existingDocEntry);

                    PersistDerivedFields(existingDocEntry);

                    LoadRecordInUpdateMode(existingDocEntry);


                    _MasterDBDataSource.SetValue("U_StartDate", 0, startDate);
                    _MasterDBDataSource.SetValue("U_EndDate", 0, endDate);

                    // Recalculate remaining days based on the potentially new dates
                    int remDays = CalculateDateDiff(startDate, endDate);
                    _MasterDBDataSource.SetValue("U_RemDays", 0, remDays.ToString());

                    // Re-check expiry status
                    CheckAndSetExpiryStatus(endDate);
                }
                else
                {
                    // No DocEntry in UI. Double check DB via Bank Facility No just in case.
                    string sqlCheck = $@"SELECT ""DocEntry"", ""U_DocNum"", ""U_TotFactAmt"", ""U_TotFctAvld"", ""U_RemFact"" 
                                 FROM ""@BANKFACTH"" 
                                 WHERE ""U_BnkFactNo"" = '{bankFacilityNo.Replace("'", "''")}' 
                                 AND ""U_BnkCode"" = '{bankCode.Replace("'", "''")}'";
                    rs.DoQuery(sqlCheck);

                    if (!rs.EoF)
                    {
                        // Record found in DB. Link it back to parent and load it.
                        string foundDocEntry = rs.Fields.Item("DocEntry").Value.ToString();
                        string totalFacility = rs.Fields.Item("U_TotFactAmt").Value.ToString();
                        string availedFacility = rs.Fields.Item("U_TotFctAvld").Value.ToString();
                        string remaining = rs.Fields.Item("U_RemFact").Value.ToString();
                        UpdateParentFormDocEntry(parentUID, parentRowIdx, foundDocEntry, totalFacility,availedFacility,remaining);
                        LoadRecordInUpdateMode(foundDocEntry);
                    }
                    else
                    {
                        // If New Run the creation logic.
                        _Form.Mode = BoFormMode.fm_ADD_MODE;

                        // 1. Set Header Basic Info
                        _MasterDBDataSource.SetValue("U_BnkCode", 0, bankCode);
                        _MasterDBDataSource.SetValue("U_BnkName", 0, bankName);
                        _MasterDBDataSource.SetValue("U_StartDate", 0, startDate);
                        _MasterDBDataSource.SetValue("U_EndDate", 0, endDate);
                        _MasterDBDataSource.SetValue("U_BnkFactNo", 0, bankFacilityNo);
                        string nextDocNum = Utilities.getMaxColumnValueNum("@BANKFACTH", "U_DocNum");
                        _MasterDBDataSource.SetValue("U_DocNum", 0, nextDocNum);
                        _MasterDBDataSource.SetValue("U_CretBy", 0, Utilities.Application.Company.UserName);

                        int remDays = CalculateDateDiff(startDate, endDate);
                        _MasterDBDataSource.SetValue("U_RemDays", 0, remDays.ToString());

                        CheckAndSetExpiryStatus(endDate);

                        // Fetch GL Account Mappings 
                        string sqlGL = $@"
                    SELECT 
                        T0.""U_TrLoan"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_TrLoan"") AS ""TrLoanName"",
                        T0.""U_StlLoan"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_StlLoan"") AS ""StlLoanName"",
                        T0.""U_DemandLoan"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_DemandLoan"") AS ""DemandLoanName"",
                        T0.""U_ODLoan"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_ODLoan"") AS ""ODLoanName"",
                        T0.""U_TermLoan"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_TermLoan"") AS ""TermLoanName"",
                        T0.""U_BridgeLoan"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_BridgeLoan"") AS ""BridgeLoanName"",
                        T0.""U_TRInterestAc"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_TRInterestAc"") AS ""TRIntName"",
                        T0.""U_STLInterestAc"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_STLInterestAc"") AS ""STLIntName"",
                        T0.""U_DemandInterestAc"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_DemandInterestAc"") AS ""DemandIntName"",
                        T0.""U_ODInterestAc"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_ODInterestAc"") AS ""ODIntName"",
                        T0.""U_TermInterestAc"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_TermInterestAc"") AS ""TermIntName"",
                        T0.""U_BridgeInterestAc"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_BridgeInterestAc"") AS ""BridgeIntName""
                    FROM ""@LOANGLAC"" T0 
                    WHERE T0.""Code"" = '{bankCode.Replace("'", "''")}'";

                        SAPbobsCOM.Recordset rsGL = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        rsGL.DoQuery(sqlGL);

                        if (!rsGL.EoF)
                        {
                            // Map Account Codes
                            _MasterDBDataSource.SetValue("U_TRLoanAct", 0, rsGL.Fields.Item("U_TrLoan").Value.ToString());
                            _MasterDBDataSource.SetValue("U_STLAcct", 0, rsGL.Fields.Item("U_StlLoan").Value.ToString());
                            _MasterDBDataSource.SetValue("U_DMLoanAct", 0, rsGL.Fields.Item("U_DemandLoan").Value.ToString());
                            _MasterDBDataSource.SetValue("U_ODAcct", 0, rsGL.Fields.Item("U_ODLoan").Value.ToString());
                            _MasterDBDataSource.SetValue("U_BrdgLAcct", 0, rsGL.Fields.Item("U_BridgeLoan").Value.ToString());
                            _MasterDBDataSource.SetValue("U_TrmLAcct", 0, rsGL.Fields.Item("U_TermLoan").Value.ToString());

                            // Map Account Names
                            _MasterDBDataSource.SetValue("U_TRLoanActN", 0, rsGL.Fields.Item("TrLoanName").Value.ToString());
                            _MasterDBDataSource.SetValue("U_STLAcctN", 0, rsGL.Fields.Item("StlLoanName").Value.ToString());
                            _MasterDBDataSource.SetValue("U_DMLoanActN", 0, rsGL.Fields.Item("DemandLoanName").Value.ToString());
                            _MasterDBDataSource.SetValue("U_ODAcctN", 0, rsGL.Fields.Item("ODLoanName").Value.ToString());
                            _MasterDBDataSource.SetValue("U_BrdgLAcctN", 0, rsGL.Fields.Item("BridgeLoanName").Value.ToString());
                            _MasterDBDataSource.SetValue("U_TrmLAcctN", 0, rsGL.Fields.Item("TermLoanName").Value.ToString());

                            // Map Interest Accounts
                            _MasterDBDataSource.SetValue("U_ITRLoanAct", 0, rsGL.Fields.Item("U_TRInterestAc").Value.ToString());
                            _MasterDBDataSource.SetValue("U_ISTLAcct", 0, rsGL.Fields.Item("U_STLInterestAc").Value.ToString());
                            _MasterDBDataSource.SetValue("U_IDMLoanAct", 0, rsGL.Fields.Item("U_DemandInterestAc").Value.ToString());
                            _MasterDBDataSource.SetValue("U_IODAcct", 0, rsGL.Fields.Item("U_ODInterestAc").Value.ToString());
                            _MasterDBDataSource.SetValue("U_IBGAcct", 0, rsGL.Fields.Item("U_BridgeInterestAc").Value.ToString());
                            _MasterDBDataSource.SetValue("U_ITRMAcct", 0, rsGL.Fields.Item("U_TermInterestAc").Value.ToString());

                            // Map Interest Names
                            _MasterDBDataSource.SetValue("U_ITRLoanActN", 0, rsGL.Fields.Item("TRIntName").Value.ToString());
                            _MasterDBDataSource.SetValue("U_ISTLAcctN", 0, rsGL.Fields.Item("STLIntName").Value.ToString());
                            _MasterDBDataSource.SetValue("U_IDMLoanActN", 0, rsGL.Fields.Item("DemandIntName").Value.ToString());
                            _MasterDBDataSource.SetValue("U_IODAcctN", 0, rsGL.Fields.Item("ODIntName").Value.ToString());
                            _MasterDBDataSource.SetValue("U_IBGAcctN", 0, rsGL.Fields.Item("BridgeIntName").Value.ToString());
                            _MasterDBDataSource.SetValue("U_ITRMAcctN", 0, rsGL.Fields.Item("TermIntName").Value.ToString());

                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(rsGL);

                        // Programmatically Add the record
                        _Form.Items.Item("1").Click();

                        string totalFacility = _MasterDBDataSource.GetValue("U_TotFactAmt", 0);
                        string totalAvialed = _MasterDBDataSource.GetValue("U_TotFctAvld", 0);
                        string remaning = _MasterDBDataSource.GetValue("U_RemFact", 0);

                        // Capture the new DocEntry and update the Parent
                        //string newDocEntry = Utilities.Application.Company.GetNewObjectKey();
                        string newDocEntry = "";
                        SAPbobsCOM.Recordset rsFetch = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        try
                        {
                            // Query by the unique facility number we just saved
                            string sqlGetId = $@"SELECT ""DocEntry"" FROM ""@BANKFACTH"" 
                                         WHERE ""U_BnkFactNo"" = '{bankFacilityNo.Replace("'", "''")}' 
                                         AND ""U_BnkCode"" = '{bankCode.Replace("'", "''")}'";
                            rsFetch.DoQuery(sqlGetId);
                            rsFetch.DoQuery(sqlGetId);

                            if (!rsFetch.EoF)
                            {
                                newDocEntry = rsFetch.Fields.Item("DocEntry").Value.ToString();
                            }
                        }

                        finally { System.Runtime.InteropServices.Marshal.ReleaseComObject(rsFetch); }

                        UpdateParentFormDocEntry(parentUID, parentRowIdx, newDocEntry, totalFacility,totalAvialed, remaning);

                        // 5. Reload this form for editing
                        LoadRecordInUpdateMode(newDocEntry);
                    }
                }
                CheckAndSetExpiryStatus(endDate);
                
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage("LoadData Error: " + ex.Message);
            }
            finally
            {
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                _Form.Freeze(false);
            }
        }
        #endregion

        #region Calculate the no of days between start and end date
        private int CalculateDateDiff(string frmDate, string toDate)
        {
            try
            {
                if (string.IsNullOrEmpty(frmDate) || string.IsNullOrEmpty(toDate)) return 0;
                DateTime start = DateTime.ParseExact(frmDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                DateTime end = DateTime.ParseExact(toDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                return (end - start).Days + 1;
            }
            catch { return 0; }
        }
        #endregion


        #region once double clicked update the docentry and other amounts on base 
        private void UpdateParentFormDocEntry(string parentUID, int rowIdx, string docEntry, string totalFacility, string totalAvailed, string remaining)
        {
            try
            {
                SAPbouiCOM.Form pForm = Utilities.Application.SBO_Application.Forms.Item(parentUID);
                SAPbouiCOM.DBDataSource pChildDS = pForm.DataSources.DBDataSources.Item("@BANKFACC1");
                SAPbouiCOM.Matrix pMatrix = (SAPbouiCOM.Matrix)pForm.Items.Item("mtxBnk").Specific;

                // Write to DataSource and refresh matrix visually
                pChildDS.SetValue("U_DocEntry", rowIdx, docEntry);
                pChildDS.SetValue("U_FacilityAmt", rowIdx, totalFacility);
                pChildDS.SetValue("U_AvailedAmt", rowIdx, totalAvailed);
                pChildDS.SetValue("U_BalanceAmt", rowIdx, remaining);

                double TotalFacility = Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("5").Specific).Value);

                if (TotalFacility > 0)
                {
                    string facilityCode = _MasterDBDataSource.GetValue("U_BnkFactNo", 0).Trim();

                    if (!string.IsNullOrEmpty(facilityCode))
                    {
                        // You must use a Recordset to run SQL in SAP B1
                        SAPbobsCOM.Recordset rsUpdate = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        try
                        {
                            //string sql = $@"UPDATE ""@BANKFACC1"" SET ""U_Status"" = 'Active' WHERE ""U_Facility"" = '{facilityCode.Replace("'", "''")}'";
                            //rsUpdate.DoQuery(sql);
                            pChildDS.SetValue("U_Status", rowIdx, "Active");
                        }
                        catch (Exception ex)
                        {
                            Utilities.ShowErrorMessage("SQL Update Error: " + ex.Message);
                        }
                        finally
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(rsUpdate);
                        }
                    }
                }

                pMatrix.LoadFromDataSource();

                // Mark parent form as 'Update' mode so the link can be saved
                if (pForm.Mode == BoFormMode.fm_OK_MODE)
                {
                    pForm.Mode = BoFormMode.fm_UPDATE_MODE;
                    pForm.Items.Item("1").Click();
                }
            }
            catch { /* Parent form might be closed */ }
        }
        #endregion

        #region Load record in Update Mode
        private void LoadRecordInUpdateMode(string docEntry)
        {
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                rs.DoQuery($"SELECT \"DocEntry\" FROM \"@BANKFACTH\" WHERE \"DocEntry\" = {docEntry}");
                if (!rs.EoF)
                {
                    //string docNum = rs.Fields.Item(0).Value.ToString();
                    string docE = rs.Fields.Item("DocEntry").Value.ToString();

                    _Form.Mode = BoFormMode.fm_FIND_MODE;
                    //_Form.Items.Item("12").Enabled = true;
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtDocE").Specific).Value = docE;
                    _Form.Items.Item("1").Click(); // Simulate 'Find' click
                }
            }
            finally { System.Runtime.InteropServices.Marshal.ReleaseComObject(rs); }
        }
        #endregion


        #region check endDate and update status accordingly
        private void CheckAndSetExpiryStatus(string endDateStr)
        {
            try
            {
                if (string.IsNullOrEmpty(endDateStr)) return;

                // SAP dates are yyyyMMdd
                DateTime endDate = DateTime.ParseExact(endDateStr, "yyyyMMdd", null);

                if (endDate < DateTime.Today)
                {
                    _MasterDBDataSource.SetValue("U_Status", 0, "Closed");
                    _Form.Items.Item("1000002").Enabled = false;
                }
                else
                {
                    // If not expired, ensure the field is enabled (unless business logic says otherwise)
                    _Form.Items.Item("1000002").Enabled = true;
                }
            }
            catch (Exception)
            {
                // Handle potential date parsing errors
            }
        }

        #endregion

        private void SyncGLAccountsFromMapping(string bankCode, string docEntry)
        {
            SAPbobsCOM.Recordset rsGL = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                // 1. Fetch Codes from Setup and Names from OACT (Subqueries ensure names are current)
                string sqlFetch = $@"
            SELECT 
                T0.""U_TrLoan"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_TrLoan"") AS ""TrN"",
                T0.""U_StlLoan"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_StlLoan"") AS ""StlN"",
                T0.""U_DemandLoan"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_DemandLoan"") AS ""DmN"",
                T0.""U_ODLoan"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_ODLoan"") AS ""OdN"",
                T0.""U_TermLoan"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_TermLoan"") AS ""TrmN"",
                T0.""U_BridgeLoan"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_BridgeLoan"") AS ""BdgN"",
                T0.""U_TRInterestAc"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_TRInterestAc"") AS ""TrIntN"",
                T0.""U_STLInterestAc"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_STLInterestAc"") AS ""StlIntN"",
                T0.""U_DemandInterestAc"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_DemandInterestAc"") AS ""DmIntN"",
                T0.""U_ODInterestAc"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_ODInterestAc"") AS ""OdIntN"",
                T0.""U_TermInterestAc"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_TermInterestAc"") AS ""TrmIntN"",
                T0.""U_BridgeInterestAc"", (SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = T0.""U_BridgeInterestAc"") AS ""BdgIntN""
            FROM ""@LOANGLAC"" T0 
            WHERE T0.""Code"" = '{bankCode.Replace("'", "''")}'";

                rsGL.DoQuery(sqlFetch);

                if (!rsGL.EoF)
                {
                    // Helper to handle nulls and escape single quotes for SQL safety
                    Func<string, string> sqlVal = (alias) => rsGL.Fields.Item(alias).Value.ToString().Replace("'", "''");

                    // 2. Perform a Bulk Update on the Header Table
                    string updateSql = $@"
                UPDATE ""@BANKFACTH"" SET 
                ""U_TRLoanAct""   = '{sqlVal("U_TrLoan")}', ""U_TRLoanActN""  = '{sqlVal("TrN")}',
                ""U_STLAcct""     = '{sqlVal("U_StlLoan")}', ""U_STLAcctN""    = '{sqlVal("StlN")}',
                ""U_DMLoanAct""   = '{sqlVal("U_DemandLoan")}', ""U_DMLoanActN""  = '{sqlVal("DmN")}',
                ""U_ODAcct""      = '{sqlVal("U_ODLoan")}', ""U_ODAcctN""      = '{sqlVal("OdN")}',
                ""U_BrdgLAcct""   = '{sqlVal("U_BridgeLoan")}', ""U_BrdgLAcctN""   = '{sqlVal("BdgN")}',
                ""U_TrmLAcct""    = '{sqlVal("U_TermLoan")}', ""U_TrmLAcctN""    = '{sqlVal("TrmN")}',
                ""U_ITRLoanAct""  = '{sqlVal("U_TRInterestAc")}', ""U_ITRLoanActN"" = '{sqlVal("TrIntN")}',
                ""U_ISTLAcct""    = '{sqlVal("U_STLInterestAc")}', ""U_ISTLAcctN""   = '{sqlVal("StlIntN")}',
                ""U_IDMLoanAct""  = '{sqlVal("U_DemandInterestAc")}', ""U_IDMLoanActN"" = '{sqlVal("DmIntN")}',
                ""U_IODAcct""     = '{sqlVal("U_ODInterestAc")}', ""U_IODAcctN""    = '{sqlVal("OdIntN")}',
                ""U_IBGAcct""     = '{sqlVal("U_BridgeInterestAc")}', ""U_IBGAcctN""     = '{sqlVal("BdgIntN")}',
                ""U_ITRMAcct""    = '{sqlVal("U_TermInterestAc")}', ""U_ITRMAcctN""   = '{sqlVal("TrmIntN")}'
                WHERE ""DocEntry"" = {docEntry}";

                    SAPbobsCOM.Recordset rsUpd = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    rsUpd.DoQuery(updateSql);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rsUpd);
                }
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.StatusBar.SetText("Error Syncing GL Mappings: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rsGL);
            }
        }

    }

}

