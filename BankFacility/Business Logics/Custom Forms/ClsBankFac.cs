using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BankFacility
{
    class ClsBankFac : Base
    {
        #region Constructor and Distructor
        public ClsBankFac() : base()
        {

        }

        ~ClsBankFac()
        {

        }
        #endregion

        #region VARIABBLE DECLARATION
        private SAPbouiCOM.DBDataSource _MasterDBDataSource, _ChildDBDataSource;
        private bool _isHandlingCFLEvent = false;
        SAPbouiCOM.ChooseFromListEvent _SysCFLEvent = null;
        string bankCode, bankName;
        public SAPbouiCOM.Matrix _Matrix;
        #endregion

        #region Form Data Event
        public override void FormData_Event(ref BusinessObjectInfo BusinessObjectInfo, ref bool oBubbleEvent)
        {
            if (BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_LOAD && !BusinessObjectInfo.BeforeAction)
            {
                if (BusinessObjectInfo.ActionSuccess)
                {
                    AddEmptyRowIfRequired();
                }
            }
        }
        #endregion

        #region Right Click event to disable (Remove)
        public override void RightClick_Event(ref ContextMenuInfo oEventInfo, ref bool oBubbleEvent)
        {
            // Disable the Remove menu (1283) before the right-click menu opens
            Utilities.Application.SBO_Application.Menus.Item("1283").Enabled = false;

        }
        #endregion

        #region Item Event
        public override void Item_Event(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            #region Before Action
            if (pVal.BeforeAction)
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1" && (_Form.Mode == BoFormMode.fm_ADD_MODE || _Form.Mode == BoFormMode.fm_UPDATE_MODE))
                        {
                            // Bank Code must not be blank
                            string bankCodeValue = _MasterDBDataSource.GetValue("U_BankCode", 0).Trim();
                            if (string.IsNullOrEmpty(bankCodeValue))
                            {
                                Utilities.Application.SBO_Application.StatusBar.SetText("Bank Code cannot be empty.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }

                            // If Facility Code exists but Status is blank, set to "Draft"
                            bool dataChanged = false;
                            _Matrix.FlushToDataSource(); // Transfer UI changes to DataSource

                            List<string> usedFacilityCodes = new List<string>();

                            for (int i = 0; i < _ChildDBDataSource.Size; i++)
                            {
                                string facility = _ChildDBDataSource.GetValue("U_Facility", i).Trim();

                                if (string.IsNullOrEmpty(facility)) continue;

                                if (usedFacilityCodes.Contains(facility))
                                {
                                    Utilities.Application.SBO_Application.StatusBar.SetText($"Row {i + 1}: Duplicate Facility Code '{facility}' detected.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    _Matrix.Columns.Item("ColFacC").Cells.Item(i + 1).Click();
                                    BubbleEvent = false;
                                    return;
                                }
                                else
                                {
                                    usedFacilityCodes.Add(facility);
                                }

                                string sDateStr = _ChildDBDataSource.GetValue("U_StartDate", i).Trim();
                                string eDateStr = _ChildDBDataSource.GetValue("U_EndDate", i).Trim();

                                if (string.IsNullOrEmpty(sDateStr))
                                {
                                    Utilities.Application.SBO_Application.StatusBar.SetText($"Row {i + 1}: Start Date is mandatory.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    _Matrix.Columns.Item("ColSDate").Cells.Item(i + 1).Click();
                                    BubbleEvent = false;
                                    return;
                                }

                                if (string.IsNullOrEmpty(eDateStr))
                                {
                                    Utilities.Application.SBO_Application.StatusBar.SetText($"Row {i + 1}: End Date is mandatory.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    _Matrix.Columns.Item("ColEDate").Cells.Item(i + 1).Click();
                                    BubbleEvent = false;
                                    return;
                                }

                                // start and end date validation
                                if (!string.IsNullOrEmpty(sDateStr) && !string.IsNullOrEmpty(eDateStr))
                                {
                                    try
                                    {
                                        DateTime startDate = DateTime.ParseExact(sDateStr, "yyyyMMdd", null);
                                        DateTime endDate = DateTime.ParseExact(eDateStr, "yyyyMMdd", null);

                                        if (startDate >= endDate)
                                        {
                                            Utilities.Application.SBO_Application.StatusBar.SetText($"Row {i + 1}: Start Date must be earlier than End Date.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                            _Matrix.Columns.Item("ColSDate").Cells.Item(i + 1).Click(); 
                                            BubbleEvent = false;
                                            return;
                                        }
                                    }
                                    catch (Exception)
                                    {
                                        // Handle parsing errors if necessary
                                    }
                                }

                                string status = _ChildDBDataSource.GetValue("U_Status", i).Trim();

                                // Check if Facility is entered but status is missing
                                if (!string.IsNullOrEmpty(facility) && string.IsNullOrEmpty(status))
                                {
                                    _ChildDBDataSource.SetValue("U_Status", i, "Draft");
                                    dataChanged = true;
                                }
                            }

                            if (dataChanged)
                            {
                                // Reload matrix so SAP recognizes the data update before the save continues
                                _Matrix.LoadFromDataSource();
                            }
                        }
                        break;

                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        if (pVal.ItemUID == "mtxBnk" && pVal.ColUID == "ColFacC" && pVal.Row > 0)
                        {
                            // Check if this specific row is already saved in the database
                            string docEntry = _ChildDBDataSource.GetValue("U_DocEntry", pVal.Row - 1).Trim();

                            if (!string.IsNullOrEmpty(docEntry) && docEntry != "0")
                            {
                                // If it exists, jump focus to the next column (Start Date)
                                // This makes the Facility Code feel "read-only"
                                _Matrix.Columns.Item("ColSDate").Cells.Item(pVal.Row).Click();

                                BubbleEvent = false; // Stop the click from actually selecting the Facility cell
                                return;
                            }
                        }
                        break;


                    case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                        if (pVal.ColUID == "#" && pVal.Row > 0)
                        {
                            BubbleEvent = false;

                            int rowIdx = pVal.Row - 1;

                            string existingDocEntry = _ChildDBDataSource.GetValue("U_DocEntry", rowIdx).Trim();

                            //string bankCode = _MasterDBDataSource.GetValue("U_BankCode", rowIdx).Trim();
                            //string bankName = _MasterDBDataSource.GetValue("U_BankName", rowIdx).Trim();
                            //string startDate = _ChildDBDataSource.GetValue("U_StartDate", rowIdx).Trim();
                            //string endDate = _ChildDBDataSource.GetValue("U_EndDate", rowIdx).Trim();
                            //string bankFacilityNo = _ChildDBDataSource.GetValue("U_Facility", rowIdx).Trim();

                            string bankCode = _MasterDBDataSource.GetValue("U_BankCode", 0).Trim();
                            string bankName = _MasterDBDataSource.GetValue("U_BankName", 0).Trim();
                            string startDate = _ChildDBDataSource.GetValue("U_StartDate", rowIdx).Trim();
                            string endDate = _ChildDBDataSource.GetValue("U_EndDate", rowIdx).Trim();
                            string bankFacilityNo = _ChildDBDataSource.GetValue("U_Facility", rowIdx).Trim();

                            ClsBankFacility objBankFact = new ClsBankFacility();
                            objBankFact.IsViewOnly = false;

                            this._Object = new ClsBankFacility();
                            Utilities.LoadForm(ref this._Object, Constants.Forms.Bank_Facility);
                            //this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                            ((ClsBankFacility)_Object).FormDefault();
                            //((ClsBankFacility)_Object).LoadData(bankCode, bankName, startDate, endDate, bankFacilityNo);
                            ((ClsBankFacility)_Object).LoadData(bankCode, bankName, startDate, endDate, bankFacilityNo, existingDocEntry, _Form.UniqueID, rowIdx);
                        }
                        break;

                        //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        //    _SysCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                        //    if(pVal.ItemUID == "20_U_E")
                        //    {
                        //        bankCode = _SysCFLEvent.SelectedObjects.GetValue("Code", 0).ToString().Trim();
                        //        bankName = _SysCFLEvent.SelectedObjects.GetValue("Name", 0).ToString().Trim();
                        //    }
                        //    break;

                }
            }
            #endregion
            #region After Action
            else
            {
                switch (pVal.EventType)
                {
                    //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                    //    _SysCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;

                    //    if (pVal.ActionSuccess && _SysCFLEvent.SelectedObjects != null && _SysCFLEvent.SelectedObjects.Rows.Count > 0)
                    //        if (pVal.ItemUID == "txtBnkC" && _SysCFLEvent.SelectedObjects != null)
                    //        {

                    //            if (_isHandlingCFLEvent) return; // Prevent re-entry

                    //            _isHandlingCFLEvent = true;
                    //            try
                    //            {
                    //                _Form.Freeze(true);

                    //                bankCode = _SysCFLEvent.SelectedObjects.GetValue("Code", 0).ToString().Trim();
                    //                bankName = _SysCFLEvent.SelectedObjects.GetValue("Name", 0).ToString().Trim();

                    //                SAPbobsCOM.Recordset rsCheck = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    //                string checkQuery = $@"SELECT COUNT(*) AS ""CNT"" FROM ""@BANKFAC"" WHERE ""U_BankCode"" = '{bankCode.Replace("'", "''")}'";
                    //                rsCheck.DoQuery(checkQuery);

                    //                int count = Convert.ToInt32(rsCheck.Fields.Item("CNT").Value);

                    //                if (count == 1)
                    //                {
                    //                    // RECORD EXISTS
                    //                    _Form.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

                    //                    //_Form.Items.Item("txtBnkC").Enabled = true;
                    //                    //((SAPbouiCOM.EditText)_Form.Items.Item("txtBnkC").Specific).Value = bankCode;

                    //                    //System.Windows.Forms.Application.DoEvents();
                    //                    //System.Threading.Thread.Sleep(50);

                    //                    //_Form.Items.Item("1").Click();

                    //                }
                    //                else
                    //                {
                    //                    _MasterDBDataSource.SetValue("U_BankCode", 0, bankCode);
                    //                    _MasterDBDataSource.SetValue("U_BankName", 0, bankName);

                    //                    _Matrix.Clear();
                    //                    _ChildDBDataSource.Clear();
                    //                    _ChildDBDataSource.InsertRecord(0);

                    //                    _Matrix.LoadFromDataSource();
                    //                    //AddSequenceNumbersToMatrix(_Matrix);
                    //                }
                    //            }
                    //            catch (Exception ex)
                    //            {
                    //                Utilities.ShowErrorMessage("Error while setting Bank Code: " + ex.Message);
                    //            }
                    //            finally
                    //            {
                    //                _Form.Freeze(false);
                    //                _isHandlingCFLEvent = false;

                    //            }
                    //        }
                    //    break;

                    case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                        // Trigger logic when user tabs out or leaves the Bank Code field
                        if (pVal.ItemUID == "txtBnkC" && pVal.ItemChanged)
                        {
                            HandleBankCodeValidation();
                        }

                        if (pVal.ItemUID == "mtxBnk" && pVal.ColUID == "ColFacC" && pVal.ActionSuccess)
                        {
                            // Only trigger if we are on the last visual row
                            if (pVal.Row == _Matrix.VisualRowCount)
                            {
                                AddEmptyRowIfRequired();
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        _SysCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                        if (pVal.ItemUID == "txtBnkC" && _SysCFLEvent.SelectedObjects != null)
                        {
                            try
                            {
                                string bankCode = _SysCFLEvent.SelectedObjects.GetValue("Code", 0).ToString().Trim();
                                ((SAPbouiCOM.EditText)_Form.Items.Item("txtBnkC").Specific).Value = bankCode;
                            }
                            catch
                            {

                            }
                        }
                            break;


                }
            }
            #endregion

        }
        #endregion

        #region Form Default
        public void FormDefault()
        {
            try
            {

                _Form.Freeze(true);

                _Form.DataBrowser.BrowseBy = "txtBnkC";
                _Form.Settings.Enabled = true;

                _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("mtxBnk").Specific;
                _MasterDBDataSource = _Form.DataSources.DBDataSources.Item("@BANKFAC");
                _ChildDBDataSource = _Form.DataSources.DBDataSources.Item("@BANKFACC1");

                //_Form.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage("Form Default Error");
            }
            finally
            {
                _Form.Freeze(false);
            }
        }
        #endregion


        #region Bank Code Validation
        private void HandleBankCodeValidation()
        {
            SAPbobsCOM.Recordset rs = null;
            if (_isHandlingCFLEvent) return;

            _isHandlingCFLEvent = true;
            try
            {
                _Form.Freeze(true);
                SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)_Form.Items.Item("txtBnkC").Specific;
                string bCode = oEdit.Value.Trim();

                if (string.IsNullOrEmpty(bCode)) return;

                rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                // Check if record exists
                string sqlCheck = $@"SELECT COUNT(*) FROM ""@BANKFAC"" WHERE ""U_BankCode"" = '{bCode.Replace("'", "''")}'";
                rs.DoQuery(sqlCheck);

                if (Convert.ToInt32(rs.Fields.Item(0).Value) > 0)
                {
                    // RECORD EXISTS -> SWITCH TO FIND MODE
                    _Form.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    _Form.Items.Item("txtBnkC").Enabled = true;
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtBnkC").Specific).Value = bCode;
                    _Form.Items.Item("1").Click();
                }
                else
                {
                    if (_Form.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        _Form.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                    string sqlName = $@"SELECT ""BankName"" FROM ODSC WHERE ""BankCode"" = '{bCode.Replace("'", "''")}'";
                    rs.DoQuery(sqlName);
                    string bName = !rs.EoF ? rs.Fields.Item(0).Value.ToString() : "";

                    _MasterDBDataSource.Clear(); // Clear any previous search data
                    _MasterDBDataSource.InsertRecord(0);
                    _MasterDBDataSource.SetValue("U_BankCode", 0, bCode);
                    _MasterDBDataSource.SetValue("U_BankName", 0, bName);

                    _Matrix.Clear();
                    _ChildDBDataSource.Clear();
                    _ChildDBDataSource.InsertRecord(0);

                    _ChildDBDataSource.SetValue("U_Facility", 0, ""); // Set an empty string to mark it 'dirty'

                    AddSequenceNumbersToMatrix();
                    //_Matrix.LoadFromDataSource();
                    _Form.Items.Item("mtxBnk").Enabled = true;

                    AddEmptyRowIfRequired();

                    _Matrix.Columns.Item("ColFacC").Cells.Item(1).Click();
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                _Form.Freeze(false);
                _isHandlingCFLEvent = false;
            }
        }
        #endregion

        #region Add Empty Row if Required
        private void AddEmptyRowIfRequired()
        {
            try
            {
                _Form.Freeze(true);
                _Matrix.FlushToDataSource();

                int rowCount = _ChildDBDataSource.Size;

                if (rowCount == 0)
                {
                    _ChildDBDataSource.InsertRecord(0);
                }
                else
                {
                    string lastFacility = _ChildDBDataSource.GetValue("U_Facility", rowCount - 1).Trim();

                    if (!string.IsNullOrEmpty(lastFacility))
                    {
                        _ChildDBDataSource.InsertRecord(rowCount);
                    }
                }

                AddSequenceNumbersToMatrix();

                //_Matrix.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                // Handle error
            }
            finally
            {
                _Form.Freeze(false);
            }
        }
        #endregion


        #region Add Sequence Numbers
        private void AddSequenceNumbersToMatrix()
        {
            try
            {
                // Update the DataSource
                for (int i = 0; i < _ChildDBDataSource.Size; i++)
                {
                    _ChildDBDataSource.SetValue("U_SrNo", i, (i + 1).ToString());
                }
                // Sync the Matrix UI with the updated DataSource
                _Matrix.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                // Handle error
            }
        }
        #endregion


    }
}

