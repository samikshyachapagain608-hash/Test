using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using SAPbobsCOM;

namespace BankFacility
{
    class ClsLoanMaster : Base
    {
        #region Constructor and Distructor
        public ClsLoanMaster() : base()
        { }

        ~ClsLoanMaster() { }
        #endregion

        #region Variable Declaration

        SAPbouiCOM.ChooseFromListEvent _SysCFLEvent = null;
        string bankCode, bankName;
        public SAPbouiCOM.DBDataSource _MasterDBDataSource = null, _ChildDBDataSource = null;
        public SAPbouiCOM.Matrix _Matrix;
        private SAPbouiCOM.Conditions _Cons;
        private SAPbouiCOM.Condition _Con;
        string query, query1, query2;
        SAPbobsCOM.Recordset _RS = null, _RS1 = null, _RS2 = null;
        private List<LoanRowData> _stashedRows = new List<LoanRowData>();
        private bool _isHandlingCFLEvent = false;
        private SAPbouiCOM.BoFormMode _lastMode;
        #endregion

        #region Form Data EVent
        public override void FormData_Event(ref BusinessObjectInfo BusinessObjectInfo, ref bool oBubbleEvent)
        {
            if (BusinessObjectInfo.BeforeAction == false && BusinessObjectInfo.ActionSuccess == true)
            {
                if (BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_LOAD)
                {
                    //_Matrix.Columns.Item("ColLoanA").Editable = false;
                    SyncCostCentersWithOPRC();
                    ////FreezeInterestColumn();
                    ///
                    //FreezeLockedRows();

                    SyncBalanceAndStatus();

                    LockRowsWithCostCenter();

                    FocusLastMatrixRow();
                }
            }
        }
        #endregion

        public override void RightClick_Event(ref ContextMenuInfo oEventInfo, ref bool oBubbleEvent)
        {
            // Disable the Remove menu (1283) before the right-click menu opens
            Utilities.Application.SBO_Application.Menus.Item("1283").Enabled = false;

        }

        #region ItemEvent
        public override void Item_Event(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            #region Before Action
            if (pVal.BeforeAction)
            {

                switch (pVal.EventType)
                {

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1")
                        {

                            if (_Form.Mode == BoFormMode.fm_ADD_MODE)
                            {
                                _Matrix.FlushToDataSource();

                                // 1. Check if there are any rows at all
                                if (_ChildDBDataSource.Size == 0)
                                {
                                    Utilities.Application.SBO_Application.StatusBar.SetText("Please add at least one loan line.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    BubbleEvent = false;
                                    return;
                                }

                                // 2. REQUIREMENT: Check if the Loan Type in the FIRST row is blank
                                string firstLoanType = _ChildDBDataSource.GetValue("U_LoanType", 0).Trim();

                                if (string.IsNullOrEmpty(firstLoanType))
                                {
                                    Utilities.Application.SBO_Application.StatusBar.SetText("First Row Error: Loan Type is required.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                                    // Highlight the specific cell to guide the user
                                    _Matrix.Columns.Item("ColLoanT").Cells.Item(1).Click();

                                    BubbleEvent = false; // STOP the Add process
                                    return;
                                }
                            }

                            if (_Form.Mode == BoFormMode.fm_ADD_MODE || _Form.Mode == BoFormMode.fm_UPDATE_MODE)
                            {
                                string bCode = _MasterDBDataSource.GetValue("U_BankCode", 0).Trim();

                                if (string.IsNullOrEmpty(bCode))
                                {
                                    Utilities.Application.SBO_Application.StatusBar.SetText("Bank Code is required. Please select a Bank.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                                    // This stops SAP from adding the document
                                    BubbleEvent = false;
                                    return;
                                }

                                _Matrix.FlushToDataSource();

                                // Validate Matrix Rows: If Loan Code exists, Bank Facility No must exist
                                for (int i = 0; i < _ChildDBDataSource.Size; i++)
                                {
                                    string loanCode = _ChildDBDataSource.GetValue("U_LoanCode", i).Trim();
                                    string loanType = _ChildDBDataSource.GetValue("U_LoanType", i).Trim();
                                    string bankFacNo = _ChildDBDataSource.GetValue("U_BankFacNo", i).Trim();
                                    string loanName = _ChildDBDataSource.GetValue("U_LoanNo", i).Trim();

                                    if (string.IsNullOrEmpty(loanType)) continue;

                                    if (string.IsNullOrEmpty(loanCode))
                                    {
                                        Utilities.Application.SBO_Application.StatusBar.SetText(
                                            $"Row {i + 1}: Loan Code is required.",
                                            SAPbouiCOM.BoMessageTime.bmt_Short,
                                            SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                                        // Highlight the problematic row/cell
                                        _Matrix.Columns.Item("ColLoanC").Cells.Item(i + 1).Click();

                                        BubbleEvent = false;
                                        return;
                                    }

                                    // If a Loan Code is generated/entered but Facility No is missing
                                    if (!string.IsNullOrEmpty(loanCode) && string.IsNullOrEmpty(bankFacNo))
                                    {
                                        Utilities.Application.SBO_Application.StatusBar.SetText(
                                            $"Row {i + 1}: Facility Code (Bank Fac No) is required for Loan {loanCode}.",
                                            SAPbouiCOM.BoMessageTime.bmt_Short,
                                            SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                                        // Highlight the problematic row/cell
                                        _Matrix.Columns.Item("ColBankF").Cells.Item(i + 1).Click();

                                        BubbleEvent = false;
                                        return;
                                    }
                                    if (!string.IsNullOrEmpty(loanCode) && string.IsNullOrEmpty(loanName))
                                    {
                                        Utilities.Application.SBO_Application.StatusBar.SetText(
                                            $"Row {i + 1}: Loan Name  is required for Loan {loanCode}.",
                                            SAPbouiCOM.BoMessageTime.bmt_Short,
                                            SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                                        // Highlight the problematic row/cell
                                        _Matrix.Columns.Item("ColLoanNo").Cells.Item(i + 1).Click();

                                        BubbleEvent = false;
                                        return;
                                    }
                                }
                            }

                            _lastMode = _Form.Mode;

                            if ((_Form.Mode == BoFormMode.fm_ADD_MODE || _Form.Mode == BoFormMode.fm_UPDATE_MODE))
                            {
                                _stashedRows.Clear();

                                // 2. CAPTURE the data from the Matrix/DBDataSource
                                _Matrix.FlushToDataSource();
                                for (int i = 0; i < _ChildDBDataSource.Size; i++)
                                {
                                    string lCode = _ChildDBDataSource.GetValue("U_LoanCode", i).Trim();
                                    if (!string.IsNullOrEmpty(lCode))
                                    {
                                        _stashedRows.Add(new LoanRowData
                                        {
                                            LoanType = _ChildDBDataSource.GetValue("U_LoanType", i).Trim(),
                                            LoanCode = lCode,
                                            LoanNo = _ChildDBDataSource.GetValue("U_LoanNo", i).Trim(),
                                            EndDate = _ChildDBDataSource.GetValue("U_EndDate", i).Trim()
                                        });
                                    }
                                }

                            }
                        }

                        //if (pVal.ItemUID == "1" && _Form.Mode == BoFormMode.fm_UPDATE_MODE)
                        //{
                        //    CreateCostCentersFromMatrix();
                        //}
                        break;


                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        if (pVal.ItemUID == "txtBnkCod" && pVal.CharPressed == 9)
                        {
                            string bCode = ((SAPbouiCOM.EditText)_Form.Items.Item("txtBnkCod").Specific).Value.Trim();
                            if (!string.IsNullOrEmpty(bCode))
                            {
                                BubbleEvent = false; 
                                FocusLastMatrixRow();
                                return;
                            }
                        }
                        break;

                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:

                        if (pVal.ItemUID == "txtBnkCod")
                        {

                            query = $@"SELECT distinct T0.""BankCode"",T0.""BankName""
                         FROM ODSC T0 
                         INNER JOIN ""@BANKFAC"" T1 ON T0.""BankCode"" = T1.""U_BankCode"" and T0.""BankName"" = T1.""U_BankName"" ";
                            Utilities.ExecuteSQL(ref _RS, query);

                            this.FilterBankCode("");
                        }

                        if (pVal.ColUID == "ColBankF")
                        {
                            this.FilterFacilityCode("");
                        }

                        if (pVal.ItemUID == "mtxLoan" && pVal.ColUID == "ColLoanT")
                        {
                            this.FilterLoanTypeByDimCode();
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                        if (pVal.ColUID == "ColSrNo" && pVal.Row > 0)
                        {
                            BubbleEvent = false;

                            int rowIdx = pVal.Row - 1;

                            string loanCode = _ChildDBDataSource.GetValue("U_LoanCode", rowIdx).Trim();
                            string facilityNo = _ChildDBDataSource.GetValue("U_BankFacNo", rowIdx).Trim();
                            string loanAmt = _ChildDBDataSource.GetValue("U_LoanAmountC", rowIdx).Trim();
                            string frmDate = _ChildDBDataSource.GetValue("U_StartDate", rowIdx).Trim();
                            string toDate = _ChildDBDataSource.GetValue("U_EndDate", rowIdx).Trim();
                            string interest = _ChildDBDataSource.GetValue("U_Interest", rowIdx).Trim();
                            string baseInterest = _ChildDBDataSource.GetValue("U_BaseInterest", rowIdx).Trim();
                            string premInterest = _ChildDBDataSource.GetValue("U_PremInterest", rowIdx).Trim();



                            ClsLoanMasterDetails detailsForm = new ClsLoanMasterDetails();
                            this._Object = detailsForm;

                            this.OpenChildForm(Constants.Forms.LOAN_DETAILS, ref pVal, true);

                            detailsForm.FormDefault();
                            detailsForm.LoadDataFromParent(loanCode, facilityNo, loanAmt, frmDate, toDate, interest, baseInterest, premInterest);
                        }
                        break;

                    case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                        if (pVal.ItemUID == "mtxLoan")
                        {
                            // 1. Only trigger if the End Date column is being validated
                            if (pVal.ColUID == "ColEndD")
                            {
                                // 2. Get the entered date directly from the matrix cell 
                                // (DataSource might not be updated yet during Validate)
                                SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)_Matrix.Columns.Item("ColEndD").Cells.Item(pVal.Row).Specific;
                                string enteredDateStr = oEdit.Value.Trim();

                                if (string.IsNullOrEmpty(enteredDateStr)) break;

                                // 3. Get the Facility Number from the matrix to use in the query
                                string facilityNo = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("ColBankF").Cells.Item(pVal.Row).Specific).Value.Trim();

                                if (string.IsNullOrEmpty(facilityNo)) break;

                                // 4. Execute the query to get the Facility's allowed End Date
                                query2 = $@"SELECT DISTINCT T0.""U_EndDate""
                         FROM ""@BANKFACC1"" T0 
                         INNER JOIN ""@BANKFAC"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" 
                         WHERE T0.""U_Facility"" = '{facilityNo}'";

                                Utilities.ExecuteSQL(ref _RS2, query2);

                                if (_RS2 != null && !_RS2.EoF)
                                {
                                    // SAP Date values from Recordset are usually DateTime or yyyyMMdd strings
                                    object rawDate = _RS2.Fields.Item("U_EndDate").Value;
                                    string limitDateStr = "";

                                    if (rawDate is DateTime dt)
                                    {
                                        limitDateStr = dt.ToString("yyyyMMdd");
                                    }
                                    else
                                    {
                                        limitDateStr = rawDate.ToString();
                                    }

                                    if (!string.IsNullOrEmpty(limitDateStr))
                                    {
                                        // 5. Compare the dates
                                        // String comparison works for yyyyMMdd (e.g., "20241231" > "20240101")
                                        if (string.Compare(enteredDateStr, limitDateStr) > 0)
                                        {
                                            // 6. Block the event and show error
                                            BubbleEvent = false;
                                            Utilities.Application.SBO_Application.StatusBar.SetText(
                                                $"Selected date ({enteredDateStr}) must be less than or equal to Facility End Date ({limitDateStr}).",
                                                SAPbouiCOM.BoMessageTime.bmt_Short,
                                                SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        }
                                    }
                                }
                            }

                            if (pVal.ColUID == "ColBankF")
                            {
                                SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)_Matrix.Columns.Item("ColBankF").Cells.Item(pVal.Row).Specific;
                                string enteredFac = oEdit.Value.Trim();

                                // If empty, we usually allow it (let the "Add" button validation handle mandatory checks)
                                if (string.IsNullOrEmpty(enteredFac)) break;

                                // Get the Bank Code from the Header
                                string bCode = _MasterDBDataSource.GetValue("U_BankCode", 0).Trim();

                                if (string.IsNullOrEmpty(bCode))
                                {
                                    BubbleEvent = false;
                                    Utilities.Application.SBO_Application.StatusBar.SetText("Please select a Bank Code first.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    return;
                                }

                                // Query to see if the typed value exists for this bank and at last check status
                                string valQuery = $@"
                                    SELECT COUNT(*) as ""Count""
                                    FROM ""@BANKFACC1"" T0 
                                    INNER JOIN ""@BANKFAC"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                                    WHERE T1.""U_BankCode"" = '{bCode.Replace("'", "''")}' 
                                    AND T0.""U_Facility"" = '{enteredFac.Replace("'", "''")}' AND T0.""U_Status"" = 'Active' ";

                                SAPbobsCOM.Recordset rsVal = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                rsVal.DoQuery(valQuery);

                                int exists = Convert.ToInt32(rsVal.Fields.Item("Count").Value);

                                if (exists == 0)
                                {
                                    // BLOCK THE USER
                                    BubbleEvent = false;
                                    Utilities.Application.SBO_Application.StatusBar.SetText(
                                        $"Invalid Facility Code '{enteredFac}' for the selected bank. Please use the selection list (Tab/CFL).",
                                        SAPbouiCOM.BoMessageTime.bmt_Short,
                                        SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                }

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(rsVal);
                            }
                        }
                        
                        break;

                    case SAPbouiCOM.BoEventTypes.et_CLICK:

                        if (pVal.ItemUID == "mtxLoan" && pVal.ColUID == "ColLoanT" && pVal.Row > 0)
                        {
                            // Check if this specific row is already saved in the database
                            string docEntry = _ChildDBDataSource.GetValue("U_CostCenterCode", pVal.Row - 1).Trim();

                            if (!string.IsNullOrEmpty(docEntry) && docEntry != "0")
                            {
                                // If it exists, jump focus to the next column (Start Date)
                                // This makes the Facility Code feel "read-only"
                                _Matrix.Columns.Item("ColLoanNo").Cells.Item(pVal.Row).Click();

                                BubbleEvent = false; // Stop the click from actually selecting the Facility cell
                                return;
                            }
                        }
                        if (pVal.ItemUID == "mtxLoan" && pVal.ColUID == "ColLoanNo" && pVal.Row > 0)
                        {
                            // Check if this specific row is already saved in the database
                            string docEntry = _ChildDBDataSource.GetValue("U_CostCenterCode", pVal.Row - 1).Trim();

                            if (!string.IsNullOrEmpty(docEntry) && docEntry != "0")
                            {
                                // If it exists, jump focus to the next column (Start Date)
                                // This makes the Facility Code feel "read-only"
                                //_Matrix.Columns.Item("ColLoanNo").Cells.Item(pVal.Row).Click();

                                BubbleEvent = false; // Stop the click from actually selecting the Facility cell
                                return;
                            }
                        }
                        if (pVal.ItemUID == "mtxLoan" && pVal.ColUID == "ColBankF" && pVal.Row > 0)
                        {
                            // Check if this specific row is already saved in the database
                            string docEntry = _ChildDBDataSource.GetValue("U_JENO", pVal.Row - 1).Trim();

                            if (!string.IsNullOrEmpty(docEntry) && docEntry != "0")
                            {
                                // If it exists, jump focus to the next column (Start Date)
                                // This makes the Facility Code feel "read-only"
                                //_Matrix.Columns.Item("ColLoanNo").Cells.Item(pVal.Row).Click();

                                BubbleEvent = false; // Stop the click from actually selecting the Facility cell
                                return;
                            }
                        }
                        break;

                    case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
                        // Check if focus is on the columns we want to protect
                        if (pVal.ItemUID == "mtxLoan" && (pVal.ColUID == "ColLoanT" || pVal.ColUID == "ColLoanNo" || pVal.ColUID == "ColBankF"))
                        {
                            if (pVal.Row > 0)
                            {
                                // Check if this row is already "committed" (has a Cost Center or JE)
                                string ccCode = _ChildDBDataSource.GetValue("U_CostCenterCode", pVal.Row - 1).Trim();
                                string jeNo = _ChildDBDataSource.GetValue("U_JENO", pVal.Row - 1).Trim();

                                if (!string.IsNullOrEmpty(ccCode) || !string.IsNullOrEmpty(jeNo))
                                {
                                    // It's a locked row. Move the focus to the first "safe" editable column (e.g., End Date)
                                    _Matrix.Columns.Item("ColEndD").Cells.Item(pVal.Row).Click();
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                        }
                        break;



                }
            }

            #endregion

            #region After Action
            else
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        //if (pVal.ItemUID == "1"  && _Form.Mode == BoFormMode.fm_ADD_MODE)
                        if (pVal.ItemUID == "1")
                        {
                            if (pVal.ActionSuccess)
                            {
                                if (_lastMode == BoFormMode.fm_ADD_MODE || _lastMode == BoFormMode.fm_UPDATE_MODE)
                                {
                                    if (_stashedRows.Count > 0)
                                    {
                                        string newKey = Utilities.Application.Company.GetNewObjectKey();

                                        // Process using our stashed list
                                        CreateCostCentersFromStash(_stashedRows);

                                        // Clear stash after processing
                                        _stashedRows.Clear();

                                        SyncCostCentersWithOPRC();
                                    }
                                }
                                _lastMode = BoFormMode.fm_OK_MODE;
                            }

                        }
                        break;

                    #region Set CFL Values
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        _SysCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;

                        //if (pVal.ItemUID == "txtBnkCod" && _SysCFLEvent.SelectedObjects != null)
                        //{

                        //    try
                        //    {
                        //        _Form.Freeze(true);

                        //        bankCode = _SysCFLEvent.SelectedObjects.GetValue("BankCode", 0).ToString().Trim();
                        //        bankName = _SysCFLEvent.SelectedObjects.GetValue("BankName", 0).ToString().Trim();
                        //        _MasterDBDataSource.SetValue("U_BankCode", 0, bankCode);
                        //        _MasterDBDataSource.SetValue("U_BankName", 0, bankName);

                        //        _Matrix.Clear();
                        //        _ChildDBDataSource.Clear();
                        //        _ChildDBDataSource.InsertRecord(0);

                        //        _Matrix.LoadFromDataSource();
                        //        AddSequenceNumbersToMatrix(_Matrix);

                        //    }
                        //    catch(Exception ex)
                        //    {

                        //    }
                        //    finally
                        //    {
                        //        _Form.Freeze(false);
                        //        System.Threading.Tasks.Task.Delay(150).ContinueWith(t =>
                        //        {
                        //            try
                        //            {
                        //                SAPbouiCOM.Cell firstCell = _Matrix.Columns.Item("ColLoanT").Cells.Item(1);
                        //                firstCell.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        //                ((SAPbouiCOM.EditText)firstCell.Specific).Active = true;
                        //            }
                        //            catch { }
                        //        });
                        //    }

                        //}

                        if (pVal.ItemUID == "txtBnkCod" && _SysCFLEvent.SelectedObjects != null)
                        {
                            if (_isHandlingCFLEvent) return; // Prevent re-entry

                            _isHandlingCFLEvent = true;
                            try
                            {
                                _Form.Freeze(true);

                                bankCode = _SysCFLEvent.SelectedObjects.GetValue("BankCode", 0).ToString().Trim();
                                bankName = _SysCFLEvent.SelectedObjects.GetValue("BankName", 0).ToString().Trim();

                                SAPbobsCOM.Recordset rsCheck = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                string checkQuery = $@"SELECT COUNT(*) AS ""CNT"" FROM ""@LOANMASTERH"" WHERE ""U_BankCode"" = '{bankCode.Replace("'", "''")}'";
                                rsCheck.DoQuery(checkQuery);

                                int count = Convert.ToInt32(rsCheck.Fields.Item("CNT").Value);

                                if (count == 1)
                                {
                                    // RECORD EXISTS
                                    _Form.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

                                    _Form.Items.Item("txtBnkCod").Enabled = true;
                                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtBnkCod").Specific).Value = bankCode;

                                    System.Windows.Forms.Application.DoEvents();
                                    System.Threading.Thread.Sleep(50);

                                    _Form.Items.Item("1").Click();

                                }
                                else
                                {
                                    _MasterDBDataSource.SetValue("U_BankCode", 0, bankCode);
                                    _MasterDBDataSource.SetValue("U_BankName", 0, bankName);

                                    _Matrix.Clear();
                                    _ChildDBDataSource.Clear();
                                    _ChildDBDataSource.InsertRecord(0);

                                    _Matrix.LoadFromDataSource();
                                    AddSequenceNumbersToMatrix(_Matrix);
                                }
                            }
                            catch (Exception ex)
                            {
                                Utilities.ShowErrorMessage("Error while setting Bank Code: " + ex.Message);
                            }
                            finally
                            {
                                _Form.Freeze(false);
                                _isHandlingCFLEvent = false;

                                // Set focus to the first cell in the matrix
                                //System.Threading.Tasks.Task.Delay(150).ContinueWith(t =>
                                //{
                                //    try
                                //    {
                                //        SAPbouiCOM.Cell firstCell = _Matrix.Columns.Item("ColLoanT").Cells.Item(1);
                                //        firstCell.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                //        ((SAPbouiCOM.EditText)firstCell.Specific).Active = true;
                                //    }
                                //    catch { }
                                //});
                            }
                        }


                        if (pVal.ItemUID == "mtxLoan" && _SysCFLEvent.SelectedObjects != null)
                        {
                            _Form.Freeze(true);
                            try
                            {
                                _Matrix.FlushToDataSource();

                                _SysCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                                string loanType = _SysCFLEvent.SelectedObjects.GetValue("Name", 0).ToString().Trim();
                                string facilityNo = _SysCFLEvent.SelectedObjects.GetValue("Code", 0).ToString().Trim();

                                int msRow = pVal.Row - 1;

                                switch (pVal.ColUID)
                                {
                                    case "ColLoanT":
                                        _ChildDBDataSource.SetValue("U_LoanType", msRow, loanType);
                                        _ChildDBDataSource.SetValue("U_Status", msRow, "Draft");
                                        //_ChildDBDataSource.SetValue("U_BalanceAmountC", msRow, "0");
                                        //string value = _ChildDBDataSource.GetValue("U_LoanType", msRow);

                                        //query1 = $@" SELECT CASE WHEN T0.""U_LoanType"" = 'TR Loan' THEN T2.""U_TrLoan"" WHEN T0.""U_LoanType"" = 'Term Loan' THEN T2.""U_TermLoan""
                                        //    WHEN T0.""U_LoanType"" = 'Bridge Gap Loan' THEN T2.""U_BridgeLoan"" WHEN T0.""U_LoanType"" = 'Demand Loan' THEN T2.""U_DemandLoan""
                                        //    WHEN T0.""U_LoanType"" = 'Over Draft Loan' THEN T2.""U_ODLoan"" WHEN T0.""U_LoanType"" = 'Short Term Loan' THEN T2.""U_StlLoan""
                                        //    END AS ""LoanGL""
                                        //FROM ""@LOANMASTERC"" T0
                                        //INNER JOIN ""@LOANMASTERH"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                                        //INNER JOIN ""@LOANGLAC"" T2 ON T1.""U_BankCode"" = T2.""Code""
                                        //WHERE '{value}' = '{loanType}' and T1.""U_BankCode"" = '{bankCode}' ";
                                        //Utilities.ExecuteSQL(ref _RS1, query1);
                                        //string glAccount = _RS1.Fields.Item("LoanGL").Value;

                                        //_ChildDBDataSource.SetValue("U_GLAccount", msRow, glAccount);


                                        try
                                        {
                                            // 2. Fetch ShortName and NextNo from the mapping table
                                            string autoCodeQuery = $@"SELECT T1.""U_ShortName"", T1.""U_NextNo"" FROM ""@LOANDC"" T1 WHERE T1.""Name"" = '{loanType.Replace("'", "''")}'";

                                            // Use a temporary recordset to fetch the code components
                                            SAPbobsCOM.Recordset rsAuto = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                            rsAuto.DoQuery(autoCodeQuery);

                                            if (!rsAuto.EoF)
                                            {
                                                string sName = rsAuto.Fields.Item("U_ShortName").Value.ToString().Trim();
                                                string nNo = rsAuto.Fields.Item("U_NextNo").Value.ToString().Trim();

                                                // 3. Concatenate the values
                                                string generatedLoanCode = sName + nNo;


                                                //added (if cost center is available donot let it create another Loan Code)
                                                string costCenter = _ChildDBDataSource.GetValue("U_CostCenterCode", msRow);
                                                if (string.IsNullOrEmpty(costCenter))
                                                {

                                                    // 4. Set the value to the U_LoanCode column in the DataSource
                                                    _ChildDBDataSource.SetValue("U_LoanCode", msRow, generatedLoanCode);
                                                }
                                            }

                                            // Clean up COM object
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(rsAuto);
                                        }
                                        catch (Exception ex)
                                        {
                                            Utilities.Application.SBO_Application.StatusBar.SetText("Error generating Loan Code: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        }


                                        string bank = _MasterDBDataSource.GetValue("U_BankCode", 0);
                                        query1 = $@"SELECT ""U_ProvisionAc"", ""U_InterestAc"" ,""U_BankGLName"",""U_BankGLCode"" FROM ""@LOANGLAC"" WHERE ""Code"" = '{bank.Replace("'", "''")}' ";
                                        Utilities.ExecuteSQL(ref _RS1, query1);
                                        string provisionAccount = _RS1.Fields.Item("U_ProvisionAc").Value;
                                        string intrestAccount = _RS1.Fields.Item("U_InterestAc").Value;
                                        string BankGLName = _RS1.Fields.Item("U_BankGLName").Value;
                                        string BankGLCode = _RS1.Fields.Item("U_BankGLCode").Value;
                                        _ChildDBDataSource.SetValue("U_ProvisionAccount", msRow, provisionAccount);
                                        _ChildDBDataSource.SetValue("U_InterestAccount", msRow, intrestAccount);
                                        _ChildDBDataSource.SetValue("U_BankGLName", msRow, BankGLName);
                                        _ChildDBDataSource.SetValue("U_BankGLAcc", msRow, BankGLCode);
                                        _ChildDBDataSource.SetValue("U_CalendarType", msRow, "Nepali");
                                        _Matrix.SetLineData(pVal.Row);

                                        int currentRowS = pVal.Row;

                                        //    if (currentRowS - 1 == _ChildDBDataSource.Size - 1)
                                        //    {
                                        //        // This is the true "new entry" row.
                                        //        // Now it's safe to add another blank row and move the cursor.
                                        //        AddNewRow();

                                        //        int nextRowToFocus = pVal.Row + 1;
                                        //        MoveFocusToNext("ColLoanT", currentRowS);
                                        //    }

                                        //MoveFocusToNext("ColLoanC", pVal.Row);
                                        break;
                                        //case "ColBankF":
                                        //    _ChildDBDataSource.SetValue("U_BankFacNo", msRow, facilityNo);
                                        //    _Matrix.SetLineData(pVal.Row);
                                        //    break;


                                }

                                _Matrix.LoadFromDataSource();
                            }
                            catch (Exception ex)
                            {
                                Utilities.ShowErrorMessage("Error while setting matrix CFL value: " + ex.Message);
                            }
                            finally
                            {
                                _Form.Freeze(false);
                            }
                        }

                        break;

                    case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                        //if (pVal.ColUID == "ColSrNo")
                        //{
                        //    this._Object = new ClsLoanMasterDetails();
                        //    this.OpenChildForm(Constants.Forms.LOAN_DETAILS, ref pVal, true);
                        //    ((ClsLoanMasterDetails)_Object).FormDefault();

                        //}
                        //if (pVal.ColUID == "ColSrNo" && pVal.Row > 0)
                        //{
                        //    int rowIdx = pVal.Row - 1;

                        //    string loanCode = _ChildDBDataSource.GetValue("U_LoanCode", rowIdx).Trim();
                        //    string facilityNo = _ChildDBDataSource.GetValue("U_BankFacNo", rowIdx).Trim();
                        //    string loanAmt = _ChildDBDataSource.GetValue("U_LoanAmount", rowIdx).Trim();

                        //    ClsLoanMasterDetails detailsForm = new ClsLoanMasterDetails();
                        //    this._Object = detailsForm;

                        //    this.OpenChildForm(Constants.Forms.LOAN_DETAILS, ref pVal, true);

                        //    detailsForm.FormDefault();
                        //    detailsForm.LoadDataFromParent(loanCode, facilityNo, loanAmt);
                        //}
                        break;

                        #endregion
                }
            }
            #endregion
        }
        #endregion

        #region FormDefault
        internal void FormDefault()
        {
            Form.Freeze(true);
            //_Form.DataBrowser.BrowseBy = "txtBnkCod";
            //_Form.Settings.Enabled = true;

            _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("mtxLoan").Specific;
            _MasterDBDataSource = _Form.DataSources.DBDataSources.Item("@LOANMASTERH");
            _ChildDBDataSource = _Form.DataSources.DBDataSources.Item("@LOANMASTERC");

            _Matrix.Columns.Item("ColLoanA").RightJustified = true;
            _Matrix.Columns.Item("ColBalAmt").RightJustified = true;

            Form.Freeze(false);
        }
        #endregion

        #region Add Sequence Number to Matrix
        private void AddSequenceNumbersToMatrix(SAPbouiCOM.Matrix oMatrix)
        {
            _ChildDBDataSource = _Form.DataSources.DBDataSources.Item("@LOANMASTERC");

            for (int i = 0; i < _ChildDBDataSource.Size; i++)
            {
                _ChildDBDataSource.SetValue("U_SrNo", i, (i + 1).ToString());
            }

            oMatrix.LoadFromDataSource();
        }
        #endregion

        #region Move Focus to Next Row
        private void MoveFocusToNext(string nextColUID, int row)
        {
            System.Threading.Tasks.Task.Delay(150).ContinueWith(t =>
            {
                try
                {
                    SAPbouiCOM.Cell nextCell = _Matrix.Columns.Item(nextColUID).Cells.Item(row);
                    nextCell.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    ((SAPbouiCOM.EditText)nextCell.Specific).Active = true;
                }
                catch { }
            });
        }
        #endregion

        #region FILTER Bank Facility Code
        internal void FilterFacilityCode(string CFLId)
        {

            _Form.Freeze(true);
            try
            {
                _Cons = new Conditions();

                _Con = _Cons.Add();
                _Con.Alias = "U_BankCode";
                _Con.Operation = BoConditionOperation.co_EQUAL;
                _Con.CondVal = bankCode;

                _Form.ChooseFromLists.Item("CFL_BCode").SetConditions(_Cons);

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

        #region FILTER Bank Code
        internal void FilterBankCode(string CFLId)
        {
            _Form.Freeze(true);
            try
            {
                SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                _RS.MoveFirst();

                for (int i = 0; i < _RS.RecordCount; i++)
                {
                    SAPbouiCOM.Condition oCon = oCons.Add();

                    oCon.BracketOpenNum = 1;
                    oCon.Alias = "BankCode";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = _RS.Fields.Item("BankCode").Value.ToString();
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                    oCon = oCons.Add();
                    oCon.Alias = "BankName";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = _RS.Fields.Item("BankName").Value.ToString();
                    oCon.BracketCloseNum = 1;

                    if (i < _RS.RecordCount - 1)
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;

                    _RS.MoveNext();
                }

                _Form.ChooseFromLists.Item("CFL_Code").SetConditions(oCons);
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

        //private void CreateCostCentersFromStash(List<LoanRowData> rowList)
        //{
        //    SAPbobsCOM.Recordset rsDim = null;
        //    SAPbobsCOM.CompanyService oCompService = null;
        //    SAPbobsCOM.ProfitCentersService oPCService = null;

        //    try
        //    {
        //        rsDim = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //        oCompService = Utilities.Application.Company.GetCompanyService();
        //        oPCService = (SAPbobsCOM.ProfitCentersService)oCompService.GetBusinessService(SAPbobsCOM.ServiceTypes.ProfitCentersService);

        //        foreach (var row in rowList)
        //        {
        //            // Check if Cost Center already exists
        //            rsDim.DoQuery($@"SELECT ""PrcCode"" FROM OPRC WHERE ""PrcCode"" = '{row.LoanCode}'");

        //            if (rsDim.EoF)
        //            {
        //                // Get Dimension Code
        //                string dimQuery = $@"SELECT T1.""U_DimCode"", T1.""U_ShortName"", T1.""U_NextNo"" FROM ""@LOANDC"" T1 WHERE T1.""Name"" = '{row.LoanType}'";
        //                rsDim.DoQuery(dimQuery);

        //                int dimensionId = 5;
        //                string concatenatedCode = "";
        //                string shortName = "";

        //                if (!rsDim.EoF)
        //                {
        //                    int.TryParse(rsDim.Fields.Item("U_DimCode").Value.ToString(), out dimensionId);

        //                    shortName = rsDim.Fields.Item("U_ShortName").Value.ToString();
        //                    string nextNo = rsDim.Fields.Item("U_NextNo").Value.ToString();

        //                    // 3. Concatenate them
        //                    concatenatedCode = shortName + nextNo;

        //                }

        //                if (string.IsNullOrEmpty(concatenatedCode))
        //                {
        //                    concatenatedCode = row.LoanCode;
        //                }

        //                // Create Profit Center
        //                SAPbobsCOM.ProfitCenter oPC = (SAPbobsCOM.ProfitCenter)oPCService.GetDataInterface(SAPbobsCOM.ProfitCentersServiceDataInterfaces.pcsProfitCenter);
        //                oPC.CenterCode = concatenatedCode;
        //                oPC.CenterName = row.LoanNo;
        //                oPC.GroupCode = rsDim.Fields.Item("U_ShortName").Value.ToString();
        //                oPC.InWhichDimension = dimensionId;
        //                oPC.Active = SAPbobsCOM.BoYesNoEnum.tYES;
        //                oPC.Effectivefrom = DateTime.Now;
        //                if (!string.IsNullOrEmpty(row.EndDate))
        //                {
        //                    try
        //                    {
        //                        oPC.EffectiveTo = DateTime.ParseExact(row.EndDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
        //                    }
        //                    catch (Exception)
        //                    {
        //                        // Fallback or log error
        //                    }
        //                }

        //                try
        //                {
        //                    oPCService.AddProfitCenter(oPC);
        //                    Utilities.Application.SBO_Application.StatusBar.SetText($"Cost Center {row.LoanCode} created.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
        //                }
        //                catch (Exception ex)
        //                {
        //                    Utilities.Application.SBO_Application.StatusBar.SetText($"Error creating {row.LoanCode}: {ex.Message}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                }
        //            }
        //            else
        //            {
        //                Utilities.Application.SBO_Application.StatusBar.SetText($"Cost Center {row.LoanCode} already exists. Skipping...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Utilities.ShowErrorMessage("CreateCostCenter Error: " + ex.Message);
        //    }
        //    finally
        //    {
        //        if (rsDim != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsDim);
        //    }
        //}

        private void CreateCostCentersFromStash(List<LoanRowData> rowList)
        {
            SAPbobsCOM.Recordset rsDim = null;
            SAPbobsCOM.CompanyService oCompService = null;
            SAPbobsCOM.ProfitCentersService oPCService = null;

            try
            {
                rsDim = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oCompService = Utilities.Application.Company.GetCompanyService();
                oPCService = (SAPbobsCOM.ProfitCentersService)oCompService.GetBusinessService(SAPbobsCOM.ServiceTypes.ProfitCentersService);

                foreach (var row in rowList)
                {
                    // Check if Cost Center already exists
                    rsDim.DoQuery($@"SELECT ""PrcCode"" FROM OPRC WHERE ""PrcCode"" = '{row.LoanCode}'");

                    if (rsDim.EoF)
                    {
                        // Get Dimension Code and Next Number
                        string dimQuery = $@"SELECT T1.""U_DimCode"", T1.""U_ShortName"", T1.""U_NextNo"" FROM ""@LOANDC"" T1 WHERE T1.""Name"" = '{row.LoanType}'";
                        rsDim.DoQuery(dimQuery);

                        int dimensionId = 5;
                        string concatenatedCode = "";
                        string shortName = "";
                        string currentNextNo = "";

                        if (!rsDim.EoF)
                        {
                            int.TryParse(rsDim.Fields.Item("U_DimCode").Value.ToString(), out dimensionId);
                            shortName = rsDim.Fields.Item("U_ShortName").Value.ToString();
                            currentNextNo = rsDim.Fields.Item("U_NextNo").Value.ToString();

                            concatenatedCode = shortName + currentNextNo;
                        }

                        if (string.IsNullOrEmpty(concatenatedCode))
                        {
                            concatenatedCode = row.LoanCode;
                        }

                        // Create Profit Center Object
                        SAPbobsCOM.ProfitCenter oPC = (SAPbobsCOM.ProfitCenter)oPCService.GetDataInterface(SAPbobsCOM.ProfitCentersServiceDataInterfaces.pcsProfitCenter);
                        oPC.CenterCode = concatenatedCode;
                        oPC.CenterName = row.LoanNo;
                        oPC.GroupCode = shortName; // Used the variable extracted above
                        oPC.InWhichDimension = dimensionId;
                        oPC.Active = SAPbobsCOM.BoYesNoEnum.tYES;
                        oPC.Effectivefrom = DateTime.Now;

                        if (!string.IsNullOrEmpty(row.EndDate))
                        {
                            try { oPC.EffectiveTo = DateTime.ParseExact(row.EndDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture); }
                            catch { /* Fallback */ }
                        }

                        try
                        {
                            // 1. Add the Profit Center
                            oPCService.AddProfitCenter(oPC);

                            // 2. SUCCESS: Now update the Next Number in @LOANDC
                            // We parse the current number, add 1, and update the table
                            if (int.TryParse(currentNextNo, out int nextNoInt))
                            {
                                int updatedNo = nextNoInt + 1;
                                // Format the number back to string (padded with zeros if necessary, e.g., .ToString("D4") for 0001)
                                string updatedNoStr = updatedNo.ToString().PadLeft(currentNextNo.Length, '0');

                                string updateSql = $@"UPDATE ""@LOANDC"" SET ""U_NextNo"" = '{updatedNoStr}' WHERE ""Name"" = '{row.LoanType}'";
                                rsDim.DoQuery(updateSql);
                            }

                            Utilities.Application.SBO_Application.StatusBar.SetText($"Cost Center {concatenatedCode} created and counter updated.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        }
                        catch (Exception ex)
                        {
                            Utilities.Application.SBO_Application.StatusBar.SetText($"Error creating {row.LoanCode}: {ex.Message}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                    }
                    else
                    {
                        Utilities.Application.SBO_Application.StatusBar.SetText($"Cost Center {row.LoanCode} already exists. Skipping...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage("CreateCostCenter Error: " + ex.Message);
            }
            finally
            {
                if (rsDim != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsDim);
            }
        }

        #region Add New Row
        private void AddNewRow()
        {
            //try
            //{
            //    _Form.Freeze(true);

            //    _Matrix.FlushToDataSource();

            //    int lastIndex = _ChildDBDataSource.Size - 1;

            //    // Check if the last row already has data before adding new one
            //    string loanType = _ChildDBDataSource.GetValue("U_LoanType", lastIndex).Trim();
            //    string loanCode = _ChildDBDataSource.GetValue("U_LoanCode", lastIndex).Trim();
            //    string loanNo = _ChildDBDataSource.GetValue("U_LoanNo", lastIndex).Trim();
            //    string bankFacNo = _ChildDBDataSource.GetValue("U_BankFacNo", lastIndex).Trim();

            //    if (!string.IsNullOrEmpty(loanType) ||
            //        !string.IsNullOrEmpty(loanCode) ||
            //        !string.IsNullOrEmpty(loanNo) ||
            //        !string.IsNullOrEmpty(bankFacNo))
            //    {
            //        // Add a new record to child data source
            //        _ChildDBDataSource.InsertRecord(_ChildDBDataSource.Size);
            //    }

            //    _Matrix.LoadFromDataSource();
            //    AddSequenceNumbersToMatrix(_Matrix);
            //}
            //catch (Exception ex)
            //{
            //    Utilities.ShowErrorMessage("Error while adding new row: " + ex.Message);
            //}
            //finally
            //{
            //    _Form.Freeze(false);
            //}

            try
            {
                _Form.Freeze(true);
                _Matrix.FlushToDataSource();

                if (_ChildDBDataSource.Size > 0)
                {
                    int lastIndex = _ChildDBDataSource.Size - 1;
                    string loanType = _ChildDBDataSource.GetValue("U_LoanType", lastIndex).Trim();
                    string costCenter = _ChildDBDataSource.GetValue("U_CostCenterCode", lastIndex).Trim();

                    // ONLY add a new row if the current last row has a Loan Type 
                    // AND it has already been processed (has a Cost Center)
                    if (!string.IsNullOrEmpty(loanType) && !string.IsNullOrEmpty(costCenter))
                    {
                        _ChildDBDataSource.InsertRecord(_ChildDBDataSource.Size);
                    }
                }
                else
                {
                    // If completely empty (new bank), add the first row
                    _ChildDBDataSource.InsertRecord(0);
                }

                _Matrix.LoadFromDataSource();
                AddSequenceNumbersToMatrix(_Matrix);
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage("Error while adding new row: " + ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
            }
        }
        #endregion

        #region Sync Cost Centers Details in Loan Master
        private void SyncCostCentersWithOPRC()
        {
            try
            {
                _Form.Freeze(true);
                bool isChanged = false;
                SAPbobsCOM.Recordset rsPC = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset rsPC1 = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                // Ensure DataSource is synced with the UI
                _Matrix.FlushToDataSource();

                for (int i = 0; i < _ChildDBDataSource.Size; i++)
                {
                    string loanCode = _ChildDBDataSource.GetValue("U_LoanCode", i).Trim();
                    string glAcc = _ChildDBDataSource.GetValue("U_GLAccount", i).Trim();

                    if (string.IsNullOrEmpty(loanCode)) continue;

                    // Query OPRC (Cost Center table) to get Code and Name
                    string sql = $@"SELECT ""PrcCode"", ""PrcName"" FROM OPRC WHERE ""PrcCode"" = '{loanCode.Replace("'", "''")}'";
                    rsPC.DoQuery(sql);

                    string sql1 = $@" SELECT T0.""TransId"",SUM(T0.""Credit"") ""TotalLoan"", T0.""RefDate"" FROM JDT1 T0 
                                   WHERE T0.""ProfitCode"" = '{loanCode.Replace("'", "''")}' OR T0.""OcrCode2"" = '{loanCode.Replace("'", "''")}'  OR T0.""OcrCode3"" = '{loanCode.Replace("'", "''")}'  
                                   OR T0.""OcrCode4"" = '{loanCode.Replace("'", "''")}'  OR T0.""OcrCode5"" = '{loanCode.Replace("'", "''")}' and T0.""Account"" = '{glAcc}'
                                    GROUP BY T0.""TransId"", T0.""RefDate"" ";
                    rsPC1.DoQuery(sql1);

                    if (!rsPC.EoF)
                    {
                        string pcCode = rsPC.Fields.Item("PrcCode").Value.ToString();
                        string pcName = rsPC.Fields.Item("PrcName").Value.ToString();

                        string transId = rsPC1.Fields.Item("TransId").Value.ToString();
                        string totaLoan = rsPC1.Fields.Item("TotalLoan").Value.ToString();
                        //string date = rsPC1.Fields.Item("refDate").Value.ToString();
                        DateTime dt = Convert.ToDateTime(rsPC1.Fields.Item("RefDate").Value);
                        string date = dt.ToString("yyyyMMdd");

                        string existingCode = _ChildDBDataSource.GetValue("U_CostCenterCode", i).Trim();
                        string existingName = _ChildDBDataSource.GetValue("U_CostCenterName", i).Trim();

                        string existingJENo = _ChildDBDataSource.GetValue("U_JENO", i).Trim();
                        string existingLoanAmt = _ChildDBDataSource.GetValue("U_LoanAmountC", i).Trim();
                        string existingStartDate = _ChildDBDataSource.GetValue("U_StartDate", i).Trim();

                        // Only update if the values are different to avoid unnecessary updates
                        if (existingCode != pcCode || existingName != pcName)
                        {
                            _ChildDBDataSource.SetValue("U_CostCenterCode", i, pcCode);
                            _ChildDBDataSource.SetValue("U_CostCenterName", i, pcName);
                            isChanged = true;
                        }
                        if ((existingLoanAmt == "0.0") && (transId != "0"))
                        {
                            //_ChildDBDataSource.SetValue("U_JENO", i, transId);
                            _ChildDBDataSource.SetValue("U_LoanAmountC", i, totaLoan);
                            _ChildDBDataSource.SetValue("U_StartDate", i, date);
                            isChanged = true;
                        }
                    }
                }

                if (isChanged)
                {
                    _Matrix.LoadFromDataSource();

                    LockRowsWithCostCenter();

                    // If we are in OK mode, changing the DataSource makes the form 'Update' mode
                    if (_Form.Mode == BoFormMode.fm_OK_MODE)
                    {
                        _Form.Mode = BoFormMode.fm_UPDATE_MODE;
                        _Form.Items.Item("1").Click(); // Clicks "Update"
                        Utilities.Application.SBO_Application.StatusBar.SetText("Cost Centers synchronized and record updated.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                    }
                }
                AddNewRow();
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage("Error syncing Cost Centers: " + ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
            }
        }
        #endregion


        //private void LockRowsWithCostCenter()
        //{
        //    try
        //    {
        //        _Form.Freeze(true);

        //        int colLoanTIdx = -1;
        //        int colLoanNoIdx = -1;
        //        int colBankFIdx = -1;

        //        // 1. Identify columns by looking at what DB Field they are bound to
        //        for (int i = 0; i < _Matrix.Columns.Count; i++)
        //        {
        //            SAPbouiCOM.Column oCol = _Matrix.Columns.Item(i);
        //            string bindAlias = oCol.DataBind.Alias;

        //            // Match based on the "alias" attribute in your XML
        //            if (bindAlias == "U_LoanType") colLoanTIdx = i + 1;
        //            if (bindAlias == "U_LoanNo") colLoanNoIdx = i + 1;
        //            if (bindAlias == "U_BankFacNo") colBankFIdx = i + 1;
        //        }

        //        // 2. Iterate through the matrix rows
        //        for (int i = 1; i <= _Matrix.VisualRowCount; i++)
        //        {
        //            // Get Cost Center from DataSource
        //            string ccCode = _ChildDBDataSource.GetValue("U_CostCenterCode", i - 1).Trim();

        //            // If Cost Center exists, isEditable should be false (Locked)
        //            bool isEditable = string.IsNullOrEmpty(ccCode);

        //            // 3. Apply the lock using the discovered indices
        //            if (colLoanTIdx != -1)
        //                _Matrix.CommonSetting.SetCellEditable(i, colLoanTIdx, isEditable);

        //            if (colLoanNoIdx != -1)
        //                _Matrix.CommonSetting.SetCellEditable(i, colLoanNoIdx, isEditable);

        //            if (colBankFIdx != -1)
        //                _Matrix.CommonSetting.SetCellEditable(i, colBankFIdx, isEditable);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        // Utilities.ShowErrorMessage(ex.Message);
        //    }
        //    finally
        //    {
        //        _Form.Freeze(false);
        //    }
        //}

        private void LockRowsWithCostCenter()
        {
            try
            {
                _Form.Freeze(true);
                _Matrix.FlushToDataSource();

                _Matrix.Columns.Item("U_LoanAmountC").Editable = false;

                // 1. Identify Column Indices (1-based for CommonSetting)
                int colLoanNoIdx = -1;
                int colBankFIdx = -1;


                for (int i = 0; i < _Matrix.Columns.Count; i++)
                {
                    string uid = _Matrix.Columns.Item(i).UniqueID;
                    if (uid == "ColLoanNo") colLoanNoIdx = i + 1;
                    if (uid == "ColBankF") colBankFIdx = i + 1;
                }

                // 2. Iterate through rows and apply locks
                for (int i = 1; i <= _Matrix.VisualRowCount; i++)
                {
                    // Skip the "empty" new row if it exists beyond datasource size
                    if (i > _ChildDBDataSource.Size) break;

                    string ccCode = _ChildDBDataSource.GetValue("U_CostCenterCode", i - 1).Trim();
                    string jeNo = _ChildDBDataSource.GetValue("U_JENO", i - 1).Trim();

                    // REQUIREMENT: If Cost Center exists OR JE No exists, row becomes read-only
                    bool isEditable = string.IsNullOrEmpty(ccCode) && string.IsNullOrEmpty(jeNo);

                    // Apply locking to specific columns
                    if (colLoanNoIdx != -1) _Matrix.CommonSetting.SetCellEditable(i, colLoanNoIdx, isEditable);
                    if (colBankFIdx != -1) _Matrix.CommonSetting.SetCellEditable(i, colBankFIdx, isEditable);

                    // You can add more columns here (Start Date, etc.)
                }
            }
            catch (Exception ex)
            {
                // Utilities.LogException(ex);
            }
            finally
            {
                _Form.Freeze(false);
            }
        }


        private void SyncBalanceAndStatus()
        {
            try
            {
                _Form.Freeze(true);
                bool isChanged = false;
                SAPbobsCOM.Recordset rsBal = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                // Ensure DataSource has the latest UI values
                _Matrix.FlushToDataSource();

                for (int i = 0; i < _ChildDBDataSource.Size; i++)                  
                {
                    string loanCode = _ChildDBDataSource.GetValue("U_LoanCode", i).Trim();
                    string jeNo = _ChildDBDataSource.GetValue("U_JENO", i).Trim();

                    if (string.IsNullOrEmpty(loanCode)) continue;

                    string sqlBal = $@"SELECT 
                            CAST(IFNULL(T0.""U_BalanceAmountC"", '0') AS DECIMAL(21,6)) AS ""CleanBalance""
                        FROM ""@LOANMASTERDETAILSH"" T0 
                        WHERE T0.""U_LoanCode"" = '{loanCode.Replace("'", "''")}'";

                    rsBal.DoQuery(sqlBal);

                    double balance = 0;
                    if (!rsBal.EoF)
                    {
                        balance = Convert.ToDouble(rsBal.Fields.Item("CleanBalance").Value);
                    }

                    // 2. Update the Balance Amount field in the Child DataSource
                    string existingBalance = _ChildDBDataSource.GetValue("U_BalanceAmountC", i).Trim();
                    if (existingBalance != balance.ToString())
                    {
                        _ChildDBDataSource.SetValue("U_BalanceAmountC", i, balance.ToString());
                        isChanged = true;
                    }


                    // 3. Status Logic: if JENo is not blank AND Balance is 0 or null, set Status to 'Closed'
                    string currentStatus = _ChildDBDataSource.GetValue("U_Status", i).Trim();
                    if (!string.IsNullOrEmpty(jeNo) && balance <= 0)
                    {
                        if (currentStatus != "Closed")
                        {
                            _ChildDBDataSource.SetValue("U_Status", i, "Closed");
                            isChanged = true;
                        }
                    }
                }

                if (isChanged)
                {
                    _Matrix.LoadFromDataSource();

                    // If the form is in OK mode, we might need to click Update to save the new balances/status
                    if (_Form.Mode == BoFormMode.fm_OK_MODE)
                    {
                        _Form.Mode = BoFormMode.fm_UPDATE_MODE;
                        _Form.Items.Item("1").Click();
                    }
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage("Error syncing balances: " + ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
            }
        }

        private void FocusLastMatrixRow()
        {
            System.Threading.Tasks.Task.Delay(100).ContinueWith(t =>
            {
                try
                {
                    int lastRow = _Matrix.VisualRowCount;
                    if (lastRow > 0)
                    {
                        _Matrix.Columns.Item("ColLoanT").Cells.Item(lastRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }
                }
                catch { }
            });
        }

        internal void FilterLoanTypeByDimCode()
        {
            try
            {
                SAPbouiCOM.ChooseFromList oCFL = _Form.ChooseFromLists.Item("CFL_LoanT");
                SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition oCon;

                // Condition 1: U_DimCode is not Null
                oCon = oCons.Add();
                oCon.Alias = "U_DimCode";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_NULL;

                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                // Condition 2: U_DimCode is not an empty string
                oCon = oCons.Add();
                oCon.Alias = "U_DimCode";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                oCon.CondVal = "";

                oCFL.SetConditions(oCons);
            }
            catch (Exception ex)
            {
                // CFL might not have been initialized yet or conditions already set
            }
        }

    }

    public class LoanRowData
    {
        public string LoanType { get; set; }
        public string LoanCode { get; set; }
        public string LoanNo { get; set; }
        public string EndDate { get; set; }

    }


}
