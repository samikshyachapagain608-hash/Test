using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BankFacility
{
    class ClsLoanMasterDetails : Base
    {
        #region Constructor and Distructor
        public ClsLoanMasterDetails() : base()
        {

        }

        ~ClsLoanMasterDetails() { }
        #endregion

        #region VARIABLE DECLARATION

        public SAPbouiCOM.DBDataSource _HeaderDS = null, _ChildDS = null, _ChildDS1 = null, _ChildDS2 = null;
        public SAPbouiCOM.Matrix _Matrix;
        public SAPbouiCOM.Matrix mtx2, mtx3;
        private bool _isProcessing = false;
        private List<dynamic> _sapPeriods = null;
        private bool _isSyncingDate = false;
        private Dictionary<string, string> _calendarCache = null;
        // Add this near your other private variables
        private Dictionary<string, DateTime> _nepaliMonthEndCache = null;
        private Dictionary<string, string> _nepaliMonthNameCache = null;

        #endregion

        #region Form Data Event
        public override void FormData_Event(ref BusinessObjectInfo BusinessObjectInfo, ref bool oBubbleEvent)
        {
            if (BusinessObjectInfo.BeforeAction == false && BusinessObjectInfo.ActionSuccess == true)
            {

                    if (BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_LOAD)
                {
                    //_Form.Freeze(true);
                    //try
                    //{
                    //    UpdatePaidAndBalanceAmounts();
                    //    UpdateLoanRepaymentMatrix();
                    //    AddEmptyRowIfRequired();
                    //}
                    //finally
                    //{
                    //    _Form.Freeze(false);
                    //}
                    AddEmptyRowIfRequired();

                    if (_isProcessing) return; // Exit if we are already in the middle of our custom logic

                    try
                    {
                        _isProcessing = true;
                        _Form.Freeze(true);

                        //UpdatePaidAndBalanceAmounts();
                        //UpdateLoanRepaymentMatrix();

                        _Form.Mode = BoFormMode.fm_UPDATE_MODE;
                        if (_Form.Mode == BoFormMode.fm_UPDATE_MODE)
                        {
                            _Form.Items.Item("1").Click(); 
                        }
                    }
                    catch (Exception ex)
                    {
                        Utilities.ShowErrorMessage(ex.Message);
                    }
                    finally
                    {
                        _Form.Freeze(false);
                        _isProcessing = false; // Reset the flag
                    }
                    _Form.Freeze(true);
                    LockBookedRows();

                    //SortDataSourceChronologically();

                    _Form.Freeze(false);

                }
            }
        }
        #endregion


        #region Right Click Event
        public override void RightClick_Event(ref ContextMenuInfo oEventInfo, ref bool oBubbleEvent)
        {
                // Disable the Remove menu (1283) before the right-click menu opens
                Utilities.Application.SBO_Application.Menus.Item("1283").Enabled = false;

        }
        #endregion


        #region Form Default
        internal void FormDefault()
        {
            _Form.Freeze(true);
            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("mtxInt").Specific;
            mtx2 = (SAPbouiCOM.Matrix)_Form.Items.Item("mtxProv").Specific;
            mtx3 = (SAPbouiCOM.Matrix)_Form.Items.Item("mtxLoanD").Specific;
            _HeaderDS = _Form.DataSources.DBDataSources.Item("@LOANMASTERDETAILSH");
            _ChildDS = _Form.DataSources.DBDataSources.Item("@LOANMASTERDETAILSC1");
            _ChildDS1 = _Form.DataSources.DBDataSources.Item("@LOANMASTERDETAILSC2");
            _ChildDS2 = _Form.DataSources.DBDataSources.Item("@LOANMASTERDETAILSC");
            _Form.Freeze(false);
        }
        #endregion

        #region Item Event
        public override void Item_Event(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            #region Before Action
            if (pVal.BeforeAction)
            {
                if (pVal.EventType == BoEventTypes.et_VALIDATE && pVal.ItemUID == "mtxInt")
                {
                    string startDateStr = _HeaderDS.GetValue("U_StartDate", 0).Trim();
                    string endDateStr = _HeaderDS.GetValue("U_EndDate", 0).Trim();

                    if (string.IsNullOrEmpty(startDateStr) || string.IsNullOrEmpty(endDateStr)) return;

                    DateTime loanStartDate = DateTime.ParseExact(startDateStr, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                    DateTime loanEndDate = DateTime.ParseExact(endDateStr, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);

                    SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)_Matrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific;
                    string enteredDateStr = oEdit.Value.Trim();
                    if (string.IsNullOrEmpty(enteredDateStr)) return;

                    
                    // VALIDATION FOR FROM DATE
                    if (pVal.ColUID == "ColFrmD")
                    {
                        DateTime enteredDate;
                        try
                        {
                            enteredDate = DateTime.ParseExact(enteredDateStr, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                        }
                        catch
                        {
                            Utilities.ShowErrorMessage("Invalid date format.");
                            BubbleEvent = false;
                            return;
                        }

                        if (pVal.Row == 1)
                        {
                            if (enteredDate < loanStartDate)
                            {
                                Utilities.ShowErrorMessage($"From Date cannot be earlier than Loan Start Date ({loanStartDate.ToShortDateString()}).");
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else
                        {
                            // Sequential validation: Current From Date must be Previous To Date + 1
                            string prevToDateStr = _ChildDS.GetValue("U_ToDate", pVal.Row - 2).Trim();
                            if (!string.IsNullOrEmpty(prevToDateStr))
                            {
                                DateTime prevToDate = DateTime.ParseExact(prevToDateStr, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                                DateTime expectedDate = prevToDate.AddDays(1);

                                if (enteredDate != expectedDate)
                                {
                                    Utilities.ShowErrorMessage($"Gap detected! From Date for row {pVal.Row} must be {expectedDate.ToShortDateString()}.");
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                        }

                        if (enteredDate > loanEndDate)
                        {
                            Utilities.ShowErrorMessage("From Date cannot be later than Loan End Date.");
                            BubbleEvent = false;
                            return;
                        }
                    }

                    // VALIDATION FOR TO DATE
                    if (pVal.ColUID == "ColToD")
                    {
                        DateTime enteredDate;
                        try
                        {
                            enteredDate = DateTime.ParseExact(enteredDateStr, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                        }
                        catch
                        {
                            Utilities.ShowErrorMessage("Invalid date format.");
                            BubbleEvent = false;
                            return;
                        }

                        // Get the current row's From Date to ensure To Date is not earlier
                        //string curFrmDateStr = _ChildDS.GetValue("U_FromDate", pVal.Row - 1).Trim();
                        SAPbouiCOM.EditText oEditFrm = (SAPbouiCOM.EditText)_Matrix.Columns.Item("ColFrmD").Cells.Item(pVal.Row).Specific;
                        string curFrmDateStr = oEditFrm.Value.Trim();
                        if (!string.IsNullOrEmpty(curFrmDateStr))
                        {
                            DateTime curFrmDate = DateTime.ParseExact(curFrmDateStr, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);

                            if (enteredDate < curFrmDate)
                            {
                                Utilities.ShowErrorMessage("To Date cannot be earlier than the From Date.");
                                BubbleEvent = false;
                                return;
                            }
                        }

                        if (enteredDate > loanEndDate)
                        {
                            Utilities.ShowErrorMessage($"To Date cannot be later than Loan End Date ({loanEndDate.ToShortDateString()}).");
                            BubbleEvent = false;
                            return;
                        }
                    }
                }

                if (pVal.ItemUID == "1" && pVal.BeforeAction)
                {
                    // Use _Form.Mode instead of pVal.FormMode for more reliability in BeforeAction
                    //if (_Form.Mode == BoFormMode.fm_ADD_MODE || _Form.Mode == BoFormMode.fm_UPDATE_MODE)
                    if (_Form.Mode == BoFormMode.fm_UPDATE_MODE)
                        {
                        if (!ValidateInterestSlabs())
                        {
                            BubbleEvent = false;
                            return;
                        }
                    }
                }
            }
            #endregion

            #region After Action
            else
            {
                if (!pVal.BeforeAction)
                {
                    if (pVal.EventType == BoEventTypes.et_VALIDATE && pVal.ItemUID == "mtxInt")
                    {
                        if (pVal.ColUID == "ColBInrst" || pVal.ColUID == "ColPInrst")
                        {
                            CalculateRowInterest(pVal.Row);
                        }
                    }

                    if (pVal.EventType == BoEventTypes.et_VALIDATE && pVal.ItemUID == "mtxInt" && pVal.ColUID == "ColToD")
                    {
                        if (pVal.ActionSuccess)
                        {
                            if (pVal.Row > 0)
                            {
                                CalculateDaysAndAddNewRow(pVal.Row);
                            }
                        }
                    }

                    if (pVal.EventType == BoEventTypes.et_VALIDATE && pVal.ItemUID == "mtxInt")
                    {
                        if (pVal.ActionSuccess)
                        {
                            // Handle both From Date and To Date columns
                            //if (!_isSyncingDate && !pVal.BeforeAction && (pVal.ColUID == "ColFrmD" || pVal.ColUID == "ColToD"))
                            if (pVal.ColUID == "ColFrmD" || pVal.ColUID == "ColToD")
                            {
                                if (pVal.Row > 0)
                                { 

                                    UpdateNepaliDateInCell(pVal.Row, pVal.ColUID);

                                    //if(pVal.ColUID == "ColToD")
                                    //{
                                    //    SAPbouiCOM.EditText oEditFrm = (SAPbouiCOM.EditText)_Matrix.Columns.Item("ColFrmD").Cells.Item(pVal.Row).Specific;
                                    //    SAPbouiCOM.EditText oEditTo = (SAPbouiCOM.EditText)_Matrix.Columns.Item("ColToD").Cells.Item(pVal.Row).Specific;

                                    //    string frmDate = oEditFrm.Value.Trim();
                                    //    string toDate = oEditTo.Value.Trim();

                                    //    if (string.IsNullOrEmpty(frmDate) || string.IsNullOrEmpty(toDate)) return;

                                    //    _Form.Freeze(true);

                                    //    int totalDays = CalculateDateDiff(frmDate, toDate);
                                    //    _ChildDS.SetValue("U_NoDays", pVal.Row - 1, totalDays.ToString());

                                    //    _Form.Freeze(false);
                                    //}
                                }
                            }
                        }
                    }
                }
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        //if (pVal.ItemUID == "1"  && _Form.Mode == BoFormMode.fm_ADD_MODE)
                        if (pVal.ItemUID == "btnCalcI" && _Form.Mode == BoFormMode.fm_OK_MODE)
                        {
                            CopyFromInterestChange();
                        }
                        if (pVal.ItemUID == "btnBookI" && _Form.Mode == BoFormMode.fm_OK_MODE)
                        {
                            CopyFromInterestChange();

                            PostInterestProvisionJE();
                        }
                        else if(pVal.ItemUID == "btnCalcI" && _Form.Mode != BoFormMode.fm_OK_MODE)
                        {
                            Utilities.ShowErrorMessage("Cannot calaculate the interest amount before document is updated.");
                        }
                        break;
                }
            }
            #endregion
        }
        #endregion

        #region Load Data from Parent Form
        public void LoadDataFromParent(string loanCode, string facilityNo, string loanAmt, string frmDate, string toDate, string interest, string baseInterest, string premInterest)
        {
            try
            {
                _Form.Freeze(true);

                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string sql = $"SELECT \"DocEntry\" FROM \"@LOANMASTERDETAILSH\" WHERE \"U_LoanCode\" = '{loanCode.Replace("'", "''")}'";
                rs.DoQuery(sql);

                if (!rs.EoF)
                {
                    // RECORD EXISTS
                    _Form.Mode = BoFormMode.fm_FIND_MODE;
                    _Form.Items.Item("txtLoanC").Enabled = true;
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtLoanC").Specific).Value = loanCode;
                    _Form.Items.Item("1").Click();
                    _Form.Items.Item("txtLoanC").Enabled = false;
                }
                else
                {
                    // RECORD DOES NOT EXIST
                    _HeaderDS.SetValue("U_LoanCode", 0, loanCode);
                    _HeaderDS.SetValue("U_FacilityNo", 0, facilityNo);
                    _HeaderDS.SetValue("U_LoanAmountC", 0, loanAmt);
                    double totalLoan = 0;
                    double.TryParse(loanAmt, out totalLoan);
                    double paidAmount = 0;
                    double balanceAmount = totalLoan - paidAmount;

                    _HeaderDS.SetValue("U_StartDate", 0, frmDate ?? "");
                    _HeaderDS.SetValue("U_EndDate", 0, toDate ?? "");

                    _HeaderDS.SetValue("U_PaidAmountC", 0, paidAmount.ToString());
                    _HeaderDS.SetValue("U_BalanceAmountC", 0, balanceAmount.ToString());

                    //DateTime fDate = DateTime.ParseExact(frmDate, "yyyyMMdd", null);
                    //DateTime tDate = DateTime.ParseExact(toDate, "yyyyMMdd", null);

                    DateTime fDate = DateTime.MinValue;
                    DateTime tDate = DateTime.MinValue;

                    bool isFromDateValid = !string.IsNullOrEmpty(frmDate) && DateTime.TryParseExact(frmDate, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out fDate);
                    bool isToDateValid = !string.IsNullOrEmpty(toDate) && DateTime.TryParseExact(toDate, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out tDate);

                    _ChildDS.Clear();
                    _ChildDS.InsertRecord(0);
                    _ChildDS.SetValue("U_FromDate", 0, frmDate);
                    _ChildDS.SetValue("U_Status", 0, "O");

                    // _ChildDS.SetValue("U_ToDate", 0, toDate);

                    string childCalculatedToDate = CalculateEndDate(frmDate);
                    _ChildDS.SetValue("U_ToDate", 0, childCalculatedToDate);

                    _ChildDS.SetValue("U_InterestPer", 0, interest);
                    _ChildDS.SetValue("U_BaseInterestPer", 0, baseInterest);
                    _ChildDS.SetValue("U_PremInterestPer", 0, premInterest);

                    //_ChildDS.SetValue("U_NFromDate", 0, GetNepaliDate(fDate));
                    //_ChildDS.SetValue("U_NToDate", 0, GetNepaliDate(tDate));




                    if (isFromDateValid)
                    {
                        _ChildDS.SetValue("U_NFromDate", 0, GetNepaliDate(fDate));
                    }

                    if (isToDateValid)
                    {
                        _ChildDS.SetValue("U_NToDate", 0, GetNepaliDate(tDate));
                    }


                    int days = CalculateDateDiff(frmDate, childCalculatedToDate);
                    _ChildDS.SetValue("U_NoDays", 0, days.ToString());

                    // If the first row we just added is complete, automatically add the second blank row
                    if (!string.IsNullOrEmpty(frmDate) && !string.IsNullOrEmpty(toDate))
                    {
                        _ChildDS.InsertRecord(1);
                    }

                    _Matrix.LoadFromDataSource();
                    AddSequenceNumbersToMatrix(_Matrix, _ChildDS);


                    if ( (!string.IsNullOrEmpty(frmDate)) && (!string.IsNullOrEmpty(toDate)) )
                    {
                        _Form.Freeze(false);
                        _Form.Items.Item("1").Click(); // Click 'Add'

                        // After adding, search for it to bring it back to the screen in OK mode
                        _Form.Mode = BoFormMode.fm_FIND_MODE;
                        _Form.Items.Item("txtLoanC").Enabled = true;
                        ((SAPbouiCOM.EditText)_Form.Items.Item("txtLoanC").Specific).Value = loanCode;
                        _Form.Items.Item("1").Click(); // Click 'Find'
                        _Form.Items.Item("txtLoanC").Enabled = false;
                    }
                    else
                    {
                        Utilities.ShowErrorMessage("Document not added: From or To Date is missing.");
                    }
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage("Error loading data: " + ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
            }
        }
        #endregion


        #region Calculate End Date
        private string CalculateEndDate(string fromDate)
        {
            SAPbobsCOM.Recordset rs = null;
            try
            {
                if (string.IsNullOrEmpty(fromDate)) return "";

                DateTime engStartDate = DateTime.ParseExact(fromDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                string nepaliStartDate = GetNepaliDateDetails(engStartDate);

                if (nepaliStartDate.Length >= 7)
                {
                    string yrMonthPrefix = nepaliStartDate.Substring(0, 7);
                    rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    string sql = $@"SELECT MAX(""U_EnglishDate"") as ""MaxEng""
                            FROM ""@NEPALICALENDER""
                            WHERE ""U_NepaliDate"" LIKE '{yrMonthPrefix}%'";
                    rs.DoQuery(sql);

                    if (!rs.EoF)
                    {
                        DateTime engLastDayOfMonth = (DateTime)rs.Fields.Item("MaxEng").Value;
                        return engLastDayOfMonth.ToString("yyyyMMdd");
                    }
                }
                return "";
            }
            catch
            {
                return "";
            }
            finally
            {
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }
        #endregion


        #region Get Nepali Date
        private string GetNepaliDateDetails(DateTime engDate)
        {
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string sql = $@"SELECT ""U_NepaliDate"" FROM ""@NEPALICALENDER"" WHERE ""U_EnglishDate"" = '{engDate:yyyyMMdd}'";
                rs.DoQuery(sql);
                if (!rs.EoF) return rs.Fields.Item(0).Value.ToString();
                return engDate.ToString("yyyy/MM/dd");
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }
        #endregion


        #region Calculate the no of days and add next row
        private void CalculateDaysAndAddNewRow(int rowIdx)
        {
            //try
            //{
            //    _Form.Freeze(true);
            //    _Matrix.FlushToDataSource();

            //    string frmDate = _ChildDS.GetValue("U_FromDate", rowIdx - 1).Trim();
            //    string toDate = _ChildDS.GetValue("U_ToDate", rowIdx - 1).Trim();
            //    //string balance = _HeaderDS.GetValue("U_BalanceAmount", 0);

            //    DateTime dFrom = DateTime.ParseExact(frmDate, "yyyyMMdd", null);
            //    DateTime dTo = DateTime.ParseExact(toDate, "yyyyMMdd", null);
            //    _ChildDS.SetValue("U_NFromDate", rowIdx - 1, GetNepaliDate(dFrom));
            //    _ChildDS.SetValue("U_NToDate", rowIdx - 1, GetNepaliDate(dTo));

            //    // Only proceed if both dates in the current row are filled
            //    if (!string.IsNullOrEmpty(frmDate) && !string.IsNullOrEmpty(toDate))
            //    {
            //        int totalDays = CalculateDateDiff(frmDate, toDate);
            //        _ChildDS.SetValue("U_NoDays", rowIdx - 1, totalDays.ToString());
            //       // _ChildDS.SetValue("U_BalanceAmount", rowIdx - 1, balance.ToString());

            //        // Add New Row only if we are currently at the last row of the matrix
            //        if (rowIdx == _Matrix.VisualRowCount)
            //        {
            //            _ChildDS.InsertRecord(_ChildDS.Size);
            //        }

            //        _Matrix.LoadFromDataSource();
            //        AddSequenceNumbersToMatrix(_Matrix, _ChildDS);

            //        _Form.Freeze(false);
            //        int nextRow = rowIdx + 1;
            //        if (nextRow <= _Matrix.VisualRowCount)
            //        {
            //            // We use a small delay or wrapping in Try/Catch to ensure 
            //            // the focus isn't stolen by the SrNo column
            //            _Matrix.Columns.Item("ColFrmD").Cells.Item(nextRow).Click(BoCellClickType.ct_Regular);
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    Utilities.ShowErrorMessage("Error: " + ex.Message);
            //}
            //finally
            //{
            //   _Form.Freeze(false);
            //}

            try
            {
                // 1. Get values directly from cells instead of Flushing the whole matrix
                SAPbouiCOM.EditText oEditFrm = (SAPbouiCOM.EditText)_Matrix.Columns.Item("ColFrmD").Cells.Item(rowIdx).Specific;
                SAPbouiCOM.EditText oEditTo = (SAPbouiCOM.EditText)_Matrix.Columns.Item("ColToD").Cells.Item(rowIdx).Specific;

                string frmDate = oEditFrm.Value.Trim();
                string toDate = oEditTo.Value.Trim();

                if (string.IsNullOrEmpty(frmDate) || string.IsNullOrEmpty(toDate)) return;

                _Form.Freeze(true);

                DateTime dFrom = DateTime.ParseExact(frmDate, "yyyyMMdd", null);
                DateTime dTo = DateTime.ParseExact(toDate, "yyyyMMdd", null);

                // 2. Update the DBDataSource for the current row
                _ChildDS.SetValue("U_FromDate", rowIdx - 1, frmDate);
                _ChildDS.SetValue("U_ToDate", rowIdx - 1, toDate);
                _ChildDS.SetValue("U_NFromDate", rowIdx - 1, GetNepaliDate(dFrom));
                _ChildDS.SetValue("U_NToDate", rowIdx - 1, GetNepaliDate(dTo));

                int totalDays = CalculateDateDiff(frmDate, toDate);
                _ChildDS.SetValue("U_NoDays", rowIdx - 1, totalDays.ToString());

                // 3. Add New Row only if we are at the last visual row
                if (rowIdx == _Matrix.VisualRowCount)
                {
                    _ChildDS.InsertRecord(_ChildDS.Size);
                    _ChildDS.SetValue("U_Status", _ChildDS.Size - 1, "O");
                    // LoadFromDataSource is ONLY needed if a new row is actually added
                    _Matrix.LoadFromDataSource();
                    AddSequenceNumbersToMatrix(_Matrix, _ChildDS);

                    // Move focus to the NEW row's From Date
                    _Matrix.Columns.Item("ColFrmD").Cells.Item(rowIdx + 1).Click();
                }
                else
                {
                    // If we didn't add a row, just update the existing row's secondary cells 
                    // visually so we don't have to call LoadFromDataSource
                    //((SAPbouiCOM.EditText)_Matrix.Columns.Item("ColNoD").Cells.Item(rowIdx).Specific).Value = totalDays.ToString();
                    //((SAPbouiCOM.EditText)_Matrix.Columns.Item("ColFrmDN").Cells.Item(rowIdx).Specific).Value = GetNepaliDate(dFrom);
                    //((SAPbouiCOM.EditText)_Matrix.Columns.Item("ColToDN").Cells.Item(rowIdx).Specific).Value = GetNepaliDate(dTo);
                    _ChildDS.Offset = rowIdx - 1;
                    _Matrix.SetLineData(rowIdx);
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage("Error: " + ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
            }
        }
        #endregion

        #region Calculate No of Days
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

        #region First Load of Loan Master Details Add Row if Required
        private void AddEmptyRowIfRequired()
        {
            try
            {
                _Form.Freeze(true);
                string loanCode = _HeaderDS.GetValue("U_LoanCode", 0);
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string sql = $"SELECT \"U_EndDate\" FROM \"@LOANMASTERC\" WHERE \"U_LoanCode\" = '{loanCode.Replace("'", "''")}'";
                rs.DoQuery(sql);
                DateTime EndDate = (DateTime)rs.Fields.Item("U_EndDate").Value;

                _HeaderDS.SetValue("U_EndDate", 0, EndDate.ToString("yyyyMMdd"));

                // Check the last row in the data source
                int lastRowIdx = _ChildDS.Size - 1;
                if (lastRowIdx < 0) { _ChildDS.InsertRecord(0); }
                else
                {
                    string fDate = _ChildDS.GetValue("U_FromDate", lastRowIdx).Trim();
                    string tDate = _ChildDS.GetValue("U_ToDate", lastRowIdx).Trim();

                    if (!string.IsNullOrEmpty(fDate) && !string.IsNullOrEmpty(tDate))
                    {
                        _ChildDS.InsertRecord(_ChildDS.Size);
                        _ChildDS.SetValue("U_Status", _ChildDS.Size - 1, "O");
                    }
                }
                _Matrix.LoadFromDataSource();
                AddSequenceNumbersToMatrix(_Matrix, _ChildDS);

                _Matrix.Columns.Item("ColFrmD").Cells.Item(lastRowIdx+2).Click(BoCellClickType.ct_Regular);
                ((SAPbouiCOM.EditText)_Matrix.Columns.Item("ColFrmD").Cells.Item(lastRowIdx + 2).Specific).Active = true;
            }
            catch { }
            finally { _Form.Freeze(false); }
        }
        #endregion

        #region When Calculate Button pressed Copy lines from Interest Change 

        //private void CopyFromInterestChange() 
        //{ 
        //    SAPbobsCOM.Recordset rsInterest = null, rsRepay = null;
        //    try {
        //        _Form.Freeze(true);
        //        rsInterest = (SAPbobsCOM.Recordset)Utilities.Application.Company .GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //        rsRepay = (SAPbobsCOM.Recordset)Utilities.Application.Company .GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //        string loanCode = _HeaderDS.GetValue("U_LoanCode", 0).Trim();

        //        if (string.IsNullOrWhiteSpace(loanCode)) return;

        //        // ========================= // CURRENT MONTH END // =========================
        //        DateTime monthEnd = new DateTime( DateTime.Today.Year, DateTime.Today.Month, DateTime.DaysInMonth( DateTime.Today.Year, DateTime.Today.Month));

        //        // ========================= // INTEREST SLABS // =========================
        //        string sqlInterest = $@" SELECT T1.""U_FromDate"", T1.""U_ToDate"", T1.""U_InterestPer"", T1.""U_NoDays"", T0.""U_BalanceAmount"" 
        //        FROM ""@LOANMASTERDETAILSH"" T0 INNER JOIN ""@LOANMASTERDETAILSC1"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" 
        //        WHERE T0.""U_LoanCode"" = '{loanCode.Replace("'", "''")}' AND T1.""U_FromDate"" IS NOT NULL ORDER BY T1.""U_FromDate"" ASC"; 
        //        rsInterest.DoQuery(sqlInterest); 

        //        // ========================= // REPAYMENT TIMELINE // First unpaid = opening // Last balance = closing // Handles multiple same-day rows // =========================
        //        string sqlRepay = $@" SELECT A.""U_PaymentDate"", A.""U_UnPaidAmount"" AS ""OpeningPrincipal"", B.""DayTotalPayment"", C.""U_BalanceAmount"" AS ""ClosingPrincipal"" 
        //        FROM ( SELECT T1.""U_PaymentDate"", T1.""U_UnPaidAmount"", ROW_NUMBER() OVER ( PARTITION BY T1.""U_PaymentDate"" ORDER BY T1.""LineId"" ASC ) AS RN 
        //        FROM ""@LOANMASTERDETAILSH"" T0 INNER JOIN ""@LOANMASTERDETAILSC"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" 
        //        WHERE T0.""U_LoanCode"" = '{loanCode.Replace("'", "''")}' AND T1.""U_PaymentDate"" IS NOT NULL )A INNER JOIN ( SELECT T1.""U_PaymentDate"", SUM(T1.""U_PaymentAmount"") AS ""DayTotalPayment"" 
        //        FROM ""@LOANMASTERDETAILSH"" T0 INNER JOIN ""@LOANMASTERDETAILSC"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" 
        //        WHERE T0.""U_LoanCode"" = '{loanCode.Replace("'", "''")}' AND T1.""U_PaymentDate"" IS NOT NULL GROUP BY T1.""U_PaymentDate"" ) B ON A.""U_PaymentDate"" = B.""U_PaymentDate"" 
        //        INNER JOIN ( SELECT T1.""U_PaymentDate"", T1.""U_BalanceAmount"", ROW_NUMBER() OVER ( PARTITION BY T1.""U_PaymentDate"" ORDER BY T1.""LineId"" DESC ) AS RN 
        //        FROM ""@LOANMASTERDETAILSH"" T0 INNER JOIN ""@LOANMASTERDETAILSC"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" 
        //        WHERE T0.""U_LoanCode"" = '{loanCode.Replace("'", "''")}' AND T1.""U_PaymentDate"" IS NOT NULL ) C ON A.""U_PaymentDate"" = C.""U_PaymentDate"" 
        //        WHERE A.RN = 1 AND C.RN = 1 ORDER BY A.""U_PaymentDate"" ASC"; 
        //        rsRepay.DoQuery(sqlRepay); 

        //        // ========================= // LOAD REPAYMENTS INTO MEMORY // =========================
        //        var repayments = new List<dynamic>();
        //        while (!rsRepay.EoF) 
        //        {
        //            repayments.Add(new {
        //                PaymentDate = Convert.ToDateTime( rsRepay.Fields.Item("U_PaymentDate").Value), 
        //                OpeningPrincipal = Convert.ToDouble( rsRepay.Fields.Item("OpeningPrincipal").Value), 
        //                TotalPayment = Convert.ToDouble( rsRepay.Fields.Item("DayTotalPayment").Value), 
        //                ClosingPrincipal = Convert.ToDouble( rsRepay.Fields.Item("ClosingPrincipal").Value) }); 
        //            rsRepay.MoveNext(); 
        //        }

        //        // ========================= // TARGET TABLE // =========================
        //        SAPbouiCOM.DBDataSource dsTarget = _Form.DataSources.DBDataSources.Item("@LOANMASTERDETAILSC2"); 
        //        dsTarget.Clear(); 
        //        int row = 0; 

        //        // ========================= // PROCESS EACH INTEREST SLAB // =========================
        //        while (!rsInterest.EoF) 
        //        { 
        //            DateTime slabFrom = Convert.ToDateTime( rsInterest.Fields.Item("U_FromDate").Value);
        //            DateTime slabTo = Convert.ToDateTime( rsInterest.Fields.Item("U_ToDate").Value);

        //            DateTime interestBookedDate = new DateTime(slabTo.Year, slabTo.Month, DateTime.DaysInMonth(slabTo.Year, slabTo.Month));

        //            double rate = Convert.ToDouble( rsInterest.Fields.Item("U_InterestPer").Value); 
        //            ///// ========================= // REPAYMENTS INSIDE CURRENT SLAB // ========================= 
        //            var slabRepayments = repayments .Where(x => x.PaymentDate >= slabFrom && x.PaymentDate <= slabTo) .OrderBy(x => x.PaymentDate) .ToList();
        //            // ========================= // DETERMINE CORRECT PRINCIPAL // 
        //            // Priority: // 1. Payment inside slab → opening principal // 2. Previous payment before slab → previous closing // 3. Future payment after slab → future opening // 4. No payment at all → header balance 
        //            // =========================
        //            double principal; 
        //            if (slabRepayments.Count > 0) 
        //            { 
        //                principal = slabRepayments.First().OpeningPrincipal; 
        //            } 
        //            else 
        //            { 
        //                var previousRepayment = repayments .Where(x => x.PaymentDate < slabFrom) .OrderByDescending(x => x.PaymentDate) .FirstOrDefault(); 
        //                if (previousRepayment != null)
        //                { 
        //                    principal = previousRepayment.ClosingPrincipal; 
        //                } 
        //                else { 
        //                    var nextRepayment = repayments .Where(x => x.PaymentDate > slabTo) .OrderBy(x => x.PaymentDate) .FirstOrDefault(); 
        //                    if (nextRepayment != null) 
        //                    { 
        //                        principal = nextRepayment.OpeningPrincipal; 
        //                    } 
        //                    else 
        //                    { 
        //                        principal = Convert.ToDouble( rsInterest.Fields.Item("U_BalanceAmount").Value);
        //                    }
        //                } 
        //            } 
        //            DateTime currentStart = slabFrom; 

        //            //========================= // NO PAYMENT INSIDE SLAB // =========================
        //            if (slabRepayments.Count == 0) 
        //            {
        //                int days = (slabTo - slabFrom).Days + 1;
        //                double interestAmt = (principal * rate / 100.0) * (days / 365.0);
        //                dsTarget.InsertRecord(row); dsTarget.SetValue( "U_FromDate", row, slabFrom.ToString("yyyyMMdd"));
        //                dsTarget.SetValue( "U_ToDate", row, slabTo.ToString("yyyyMMdd")); 
        //                dsTarget.SetValue( "U_InterestPer", row, rate.ToString()); 
        //                dsTarget.SetValue( "U_NoDays", row, days.ToString()); 
        //                dsTarget.SetValue( "U_PrincipalAmt", row, principal.ToString());
        //                dsTarget.SetValue( "U_InterestAmt", row, interestAmt.ToString());
        //                dsTarget.SetValue("U_IntrDateBook",row,interestBookedDate.ToString("yyyyMMdd"));
        //                dsTarget.SetValue("U_Remarks",row, "Due to Interest Period.");

        //                row++; 
        //            } 
        //            else 
        //            { 
        //                // ========================= // SPLIT BY PAYMENT EVENTS // =========================
        //                foreach (var pay in slabRepayments) 
        //                { 
        //                    // Before payment
        //                    DateTime segmentEnd = pay.PaymentDate.AddDays(-1); 
        //                    if (segmentEnd >= currentStart) 
        //                    { 
        //                        int days = (segmentEnd - currentStart).Days + 1;
        //                        double interestAmt = (principal * rate / 100.0) * (days / 365.0);
        //                        dsTarget.InsertRecord(row);
        //                        dsTarget.SetValue( "U_FromDate", row, currentStart.ToString("yyyyMMdd"));
        //                        dsTarget.SetValue( "U_ToDate", row, segmentEnd.ToString("yyyyMMdd"));
        //                        dsTarget.SetValue( "U_InterestPer", row, rate.ToString());
        //                        dsTarget.SetValue( "U_NoDays", row, days.ToString());
        //                        dsTarget.SetValue( "U_PrincipalAmt", row, principal.ToString());
        //                        dsTarget.SetValue( "U_InterestAmt", row, interestAmt.ToString());
        //                        dsTarget.SetValue("U_IntrDateBook", row, interestBookedDate.ToString("yyyyMMdd"));
        //                        dsTarget.SetValue("U_Remarks", row, "Interest Before Repayment on " + pay.PaymentDate.ToString("dd-MM-yyyy"));

        //                        row++; 
        //                    } 
        //                    // Payment day onward
        //                    principal = pay.ClosingPrincipal; 
        //                    currentStart = pay.PaymentDate; 
        //                } 
        //                // ========================= // FINAL SEGMENT AFTER LAST PAYMENT // =========================
        //                if (currentStart <= slabTo) 
        //                {
        //                    int days = (slabTo - currentStart).Days + 1;
        //                    double interestAmt = (principal * rate / 100.0) * (days / 365.0);
        //                    DateTime lastPaymentDate = slabRepayments.Last().PaymentDate;
        //                    dsTarget.InsertRecord(row); 
        //                    dsTarget.SetValue( "U_FromDate", row, currentStart.ToString("yyyyMMdd"));
        //                    dsTarget.SetValue( "U_ToDate", row, slabTo.ToString("yyyyMMdd"));
        //                    dsTarget.SetValue( "U_InterestPer", row, rate.ToString());
        //                    dsTarget.SetValue( "U_NoDays", row, days.ToString());
        //                    dsTarget.SetValue( "U_PrincipalAmt", row, principal.ToString());
        //                    dsTarget.SetValue( "U_InterestAmt", row, interestAmt.ToString());
        //                    dsTarget.SetValue("U_IntrDateBook", row, interestBookedDate.ToString("yyyyMMdd"));
        //                    dsTarget.SetValue("U_Remarks", row, "Interest After Repayment on " + lastPaymentDate.ToString("dd-MM-yyyy"));

        //                    row++; 
        //                } 
        //            } 
        //            rsInterest.MoveNext(); 
        //        } 
        //        // ========================= // LOAD TO MATRIX // =========================
        //        mtx2 = (SAPbouiCOM.Matrix)_Form.Items.Item("mtxProv").Specific; mtx2.LoadFromDataSource();
        //        AddSequenceNumbersToMatrix(mtx2, _ChildDS1);
        //        _Form.Mode = BoFormMode.fm_UPDATE_MODE;
        //        //_Form.Items.Item("1").Click();
        //    } 
        //    catch (Exception ex)
        //    { 
        //        Utilities.ShowErrorMessage( "Query Error: " + ex.Message); 
        //    } 
        //    finally 
        //    { 
        //        if (rsInterest != null) System.Runtime.InteropServices.Marshal.ReleaseComObject( rsInterest);
        //        if (rsRepay != null) System.Runtime.InteropServices.Marshal.ReleaseComObject( rsRepay);
        //        _Form.Freeze(false);
        //    }
        //} 

        private void CopyFromInterestChange1()
        {
            SAPbobsCOM.Recordset rsInterest = null, rsRepay = null, rsExisting = null;

            try
            {
                _Form.Freeze(true);

                rsInterest = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rsRepay = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rsExisting = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string loanCode = _HeaderDS.GetValue("U_LoanCode", 0).Trim();
                if (string.IsNullOrWhiteSpace(loanCode)) return;

                // ==========================================
                // 1. FETCH EXISTING BOOKED ENTRIES FROM DB
                // ==========================================
                // We need to know which periods already have a Journal Entry (U_JENo)
                // Adjust "U_JENo" to your actual Journal Entry field name.
                var bookedEntries = new List<dynamic>();
                string sqlCheckExisting = $@"SELECT ""U_FromDate"", ""U_ToDate"", ""U_InterestAmt"", ""U_JENo"" 
                                    FROM ""@LOANMASTERDETAILSC2"" 
                                    WHERE ""DocEntry"" = (SELECT ""DocEntry"" FROM ""@LOANMASTERDETAILSH"" WHERE ""U_LoanCode"" = '{loanCode.Replace("'", "''")}')
                                    AND IFNULL(""U_JENo"", '') <> ''";

                rsExisting.DoQuery(sqlCheckExisting);
                while (!rsExisting.EoF)
                {
                    bookedEntries.Add(new
                    {
                        FromDate = Convert.ToDateTime(rsExisting.Fields.Item("U_FromDate").Value),
                        ToDate = Convert.ToDateTime(rsExisting.Fields.Item("U_ToDate").Value),
                        InterestAmt = Convert.ToDouble(rsExisting.Fields.Item("U_InterestAmt").Value),
                        JENo = rsExisting.Fields.Item("U_JENo").Value.ToString()
                    });
                    rsExisting.MoveNext();
                }

                // =========================
                // INTEREST SLABS & REPAYMENTS (Existing Logic)
                // =========================
                string sqlInterest = $@"SELECT T1.""U_FromDate"", T1.""U_ToDate"", T1.""U_InterestPer"", T1.""U_NoDays"", T0.""U_BalanceAmount""
                                FROM ""@LOANMASTERDETAILSH"" T0
                                INNER JOIN ""@LOANMASTERDETAILSC1"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                                WHERE T0.""U_LoanCode"" = '{loanCode.Replace("'", "''")}' AND T1.""U_FromDate"" IS NOT NULL
                                ORDER BY T1.""U_FromDate"" ASC";
                rsInterest.DoQuery(sqlInterest);

                string sqlRepay = $@"SELECT A.""U_PaymentDate"", A.""U_UnPaidAmountC"" AS ""OpeningPrincipal"", B.""DayTotalPayment"", C.""U_BalanceAmount"" AS ""ClosingPrincipal""
                            FROM (SELECT T1.""U_PaymentDate"", T1.""U_UnPaidAmountC"", ROW_NUMBER() OVER (PARTITION BY T1.""U_PaymentDate"" ORDER BY T1.""LineId"" ASC) AS RN
                                  FROM ""@LOANMASTERDETAILSH"" T0 INNER JOIN ""@LOANMASTERDETAILSC"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                                  WHERE T0.""U_LoanCode"" = '{loanCode.Replace("'", "''")}' AND T1.""U_PaymentDate"" IS NOT NULL) A
                            INNER JOIN (SELECT T1.""U_PaymentDate"", SUM(T1.""U_PaymentAmount"") AS ""DayTotalPayment""
                                        FROM ""@LOANMASTERDETAILSH"" T0 INNER JOIN ""@LOANMASTERDETAILSC"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                                        WHERE T0.""U_LoanCode"" = '{loanCode.Replace("'", "''")}' AND T1.""U_PaymentDate"" IS NOT NULL GROUP BY T1.""U_PaymentDate"") B ON A.""U_PaymentDate"" = B.""U_PaymentDate""
                            INNER JOIN (SELECT T1.""U_PaymentDate"", T1.""U_BalanceAmount"", ROW_NUMBER() OVER (PARTITION BY T1.""U_PaymentDate"" ORDER BY T1.""LineId"" DESC) AS RN
                                        FROM ""@LOANMASTERDETAILSH"" T0 INNER JOIN ""@LOANMASTERDETAILSC"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                                        WHERE T0.""U_LoanCode"" = '{loanCode.Replace("'", "''")}' AND T1.""U_PaymentDate"" IS NOT NULL) C ON A.""U_PaymentDate"" = C.""U_PaymentDate""
                            WHERE A.RN = 1 AND C.RN = 1 ORDER BY A.""U_PaymentDate"" ASC";
                rsRepay.DoQuery(sqlRepay);

                var repayments = new List<dynamic>();
                while (!rsRepay.EoF)
                {
                    repayments.Add(new
                    {
                        PaymentDate = Convert.ToDateTime(rsRepay.Fields.Item("U_PaymentDate").Value),
                        OpeningPrincipal = Convert.ToDouble(rsRepay.Fields.Item("OpeningPrincipal").Value),
                        TotalPayment = Convert.ToDouble(rsRepay.Fields.Item("DayTotalPayment").Value),
                        ClosingPrincipal = Convert.ToDouble(rsRepay.Fields.Item("ClosingPrincipal").Value)
                    });
                    rsRepay.MoveNext();
                }

                SAPbouiCOM.DBDataSource dsTarget = _Form.DataSources.DBDataSources.Item("@LOANMASTERDETAILSC2");
                dsTarget.Clear();
                int row = 0;
                DateTime? previousBookedMonth = null;
                double monthlyInterestTotal = 0;

                while (!rsInterest.EoF)
                {
                    DateTime slabFrom = Convert.ToDateTime(rsInterest.Fields.Item("U_FromDate").Value);
                    DateTime slabTo = Convert.ToDateTime(rsInterest.Fields.Item("U_ToDate").Value);
                    DateTime interestBookedDate = new DateTime(slabTo.Year, slabTo.Month, DateTime.DaysInMonth(slabTo.Year, slabTo.Month));
                    double rate = Convert.ToDouble(rsInterest.Fields.Item("U_InterestPer").Value);

                    var slabRepayments = repayments.Where(x => x.PaymentDate >= slabFrom && x.PaymentDate <= slabTo).OrderBy(x => x.PaymentDate).ToList();
                    double principal;

                    if (slabRepayments.Count > 0) principal = slabRepayments.First().OpeningPrincipal;
                    else
                    {
                        var previousRepayment = repayments.Where(x => x.PaymentDate < slabFrom).OrderByDescending(x => x.PaymentDate).FirstOrDefault();
                        if (previousRepayment != null) principal = previousRepayment.ClosingPrincipal;
                        else
                        {
                            var nextRepayment = repayments.Where(x => x.PaymentDate > slabTo).OrderBy(x => x.PaymentDate).FirstOrDefault();
                            principal = (nextRepayment != null) ? nextRepayment.OpeningPrincipal : Convert.ToDouble(rsInterest.Fields.Item("U_BalanceAmount").Value);
                        }
                    }

                    DateTime currentStart = slabFrom;

                    // ==========================================
                    // LOGIC HELPER: Process segment calculation
                    // ==========================================
                    Action<DateTime, DateTime, string> ProcessSegment = (start, end, remark) =>
                    {
                        // 1. Skip monthly total rows (logic handles that outside)
                        // 2. Check if this specific period already has a JE in DB
                        var existingJE = bookedEntries.FirstOrDefault(e => e.FromDate == start && e.ToDate == end);

                        // Insert monthly summary if month changed
                        if (previousBookedMonth.HasValue && previousBookedMonth.Value != interestBookedDate)
                        {
                            //AddSummaryRow(dsTarget, ref row, previousBookedMonth.Value, ref monthlyInterestTotal);
                        }
                        previousBookedMonth = interestBookedDate;

                        int days = (end - start).Days + 1;
                        double interestAmt = 0;

                        string finalRemark = remark;

                        if (existingJE != null)
                        {
                            // If JE exists, DO NOT CALCULATE. Use existing amount.
                            interestAmt = existingJE.InterestAmt;
                            finalRemark = remark;
                        }
                        else
                        {
                            // Calculate normally
                            interestAmt = (principal * rate / 100.0) * (days / 365.0);
                        }

                        dsTarget.InsertRecord(row);
                        dsTarget.SetValue("U_FromDate", row, start.ToString("yyyyMMdd"));
                        dsTarget.SetValue("U_ToDate", row, end.ToString("yyyyMMdd"));
                        dsTarget.SetValue("U_InterestPer", row, rate.ToString());
                        dsTarget.SetValue("U_NoDays", row, days.ToString());
                        dsTarget.SetValue("U_PrincipalAmtC", row, principal.ToString());
                        dsTarget.SetValue("U_InterestAmtC", row, interestAmt.ToString());
                        dsTarget.SetValue("U_IntrDateBook", row, interestBookedDate.ToString("yyyyMMdd"));
                        dsTarget.SetValue("U_Remarks", row, finalRemark);

                        // If it has a JE, preserve it in the datasource so it stays locked/referenced
                        if (existingJE != null)
                            dsTarget.SetValue("U_JENo", row, existingJE.JENo);

                        monthlyInterestTotal += interestAmt;
                        row++;
                    };

                    // ==========================================
                    // EXECUTE SEGMENTS
                    // ==========================================
                    if (slabRepayments.Count == 0)
                    {
                        ProcessSegment(slabFrom, slabTo, "Due to Interest Period.");
                    }
                    else
                    {
                        foreach (var pay in slabRepayments)
                        {
                            DateTime segmentEnd = pay.PaymentDate.AddDays(-1);
                            if (segmentEnd >= currentStart)
                            {
                                ProcessSegment(currentStart, segmentEnd, "Interest Before Repayment on " + pay.PaymentDate.ToString("dd-MM-yyyy"));
                            }
                            principal = pay.ClosingPrincipal;
                            currentStart = pay.PaymentDate;
                        }

                        if (currentStart <= slabTo)
                        {
                            ProcessSegment(currentStart, slabTo, "Interest After Repayment on " + slabRepayments.Last().PaymentDate.ToString("dd-MM-yyyy"));
                        }
                    }

                    rsInterest.MoveNext();
                }

                // Final summary row
                if (previousBookedMonth.HasValue)
                {
                    //AddSummaryRow(dsTarget, ref row, previousBookedMonth.Value, ref monthlyInterestTotal);
                }

                mtx2 = (SAPbouiCOM.Matrix)_Form.Items.Item("mtxProv").Specific;
                mtx2.LoadFromDataSource();
                AddSequenceNumbersToMatrix(mtx2, _ChildDS1);

                _Form.Mode = BoFormMode.fm_UPDATE_MODE;
                _Form.Items.Item("1").Click();
                LockBookedRows();
            }
            catch (Exception ex) { Utilities.ShowErrorMessage("Query Error: " + ex.Message); }
            finally
            {
                if (rsInterest != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsInterest);
                if (rsRepay != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsRepay);
                if (rsExisting != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsExisting);
                _Form.Freeze(false);
            }
        }

        private void CopyFromInterestChange2()
        {
            SAPbobsCOM.Recordset rsInterest = null, rsRepay = null, rsExisting = null;

            try
            {
                _Form.Freeze(true);

                rsInterest = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rsRepay = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rsExisting = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string loanCode = _HeaderDS.GetValue("U_LoanCode", 0).Trim();
                if (string.IsNullOrWhiteSpace(loanCode)) return;

                // 1. FETCH ALREADY POSTED ENTRIES (To prevent recalculating rows with JEs)
                var bookedEntries = new List<dynamic>(); // <--- DEFINITION WAS MISSING
                string sqlCheckExisting = $@"SELECT ""U_FromDate"", ""U_ToDate"", ""U_InterestAmt"", ""U_JENo"" 
                            FROM ""@LOANMASTERDETAILSC2"" 
                            WHERE ""DocEntry"" = (SELECT ""DocEntry"" FROM ""@LOANMASTERDETAILSH"" WHERE ""U_LoanCode"" = '{loanCode.Replace("'", "''")}')
                            AND IFNULL(""U_JENo"", '') <> ''";

                rsExisting.DoQuery(sqlCheckExisting);
                while (!rsExisting.EoF)
                {
                    bookedEntries.Add(new
                    {
                        FromDate = Convert.ToDateTime(rsExisting.Fields.Item("U_FromDate").Value),
                        ToDate = Convert.ToDateTime(rsExisting.Fields.Item("U_ToDate").Value),
                        InterestAmt = Convert.ToDouble(rsExisting.Fields.Item("U_InterestAmt").Value),
                        JENo = rsExisting.Fields.Item("U_JENo").Value.ToString()
                    });
                    rsExisting.MoveNext();
                }

                // 2. FETCH REPAYMENTS
                string sqlRepay = $@"SELECT A.""U_PaymentDate"", A.""U_UnPaidAmountC"" AS ""OpeningPrincipal"", B.""DayTotalPayment"", C.""U_BalanceAmount"" AS ""ClosingPrincipal""
                    FROM (SELECT T1.""U_PaymentDate"", T1.""U_UnPaidAmountC"", ROW_NUMBER() OVER (PARTITION BY T1.""U_PaymentDate"" ORDER BY T1.""LineId"" ASC) AS RN
                          FROM ""@LOANMASTERDETAILSH"" T0 INNER JOIN ""@LOANMASTERDETAILSC"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                          WHERE T0.""U_LoanCode"" = '{loanCode.Replace("'", "''")}' AND T1.""U_PaymentDate"" IS NOT NULL) A
                    INNER JOIN (SELECT T1.""U_PaymentDate"", SUM(T1.""U_PaymentAmount"") AS ""DayTotalPayment""
                                FROM ""@LOANMASTERDETAILSH"" T0 INNER JOIN ""@LOANMASTERDETAILSC"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                                WHERE T0.""U_LoanCode"" = '{loanCode.Replace("'", "''")}' AND T1.""U_PaymentDate"" IS NOT NULL GROUP BY T1.""U_PaymentDate"") B ON A.""U_PaymentDate"" = B.""U_PaymentDate""
                    INNER JOIN (SELECT T1.""U_PaymentDate"", T1.""U_BalanceAmount"", ROW_NUMBER() OVER (PARTITION BY T1.""U_PaymentDate"" ORDER BY T1.""LineId"" DESC) AS RN
                                FROM ""@LOANMASTERDETAILSH"" T0 INNER JOIN ""@LOANMASTERDETAILSC"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                                WHERE T0.""U_LoanCode"" = '{loanCode.Replace("'", "''")}' AND T1.""U_PaymentDate"" IS NOT NULL) C ON A.""U_PaymentDate"" = C.""U_PaymentDate""
                    WHERE A.RN = 1 AND C.RN = 1 ORDER BY A.""U_PaymentDate"" ASC";
                rsRepay.DoQuery(sqlRepay);

                var repayments = new List<dynamic>();
                while (!rsRepay.EoF)
                {
                    repayments.Add(new
                    {
                        PaymentDate = Convert.ToDateTime(rsRepay.Fields.Item("U_PaymentDate").Value),
                        OpeningPrincipal = Convert.ToDouble(rsRepay.Fields.Item("OpeningPrincipal").Value),
                        ClosingPrincipal = Convert.ToDouble(rsRepay.Fields.Item("ClosingPrincipal").Value)
                    });
                    rsRepay.MoveNext();
                }

                // 3. FETCH INTEREST RATE SLABS
                string sqlInterest = $@"SELECT T1.""U_FromDate"", T1.""U_ToDate"", T1.""U_InterestPer"", T0.""U_BalanceAmount""
                        FROM ""@LOANMASTERDETAILSH"" T0
                        INNER JOIN ""@LOANMASTERDETAILSC1"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                        WHERE T0.""U_LoanCode"" = '{loanCode.Replace("'", "''")}' AND T1.""U_FromDate"" IS NOT NULL
                        ORDER BY T1.""U_FromDate"" ASC";
                rsInterest.DoQuery(sqlInterest);

                SAPbouiCOM.DBDataSource dsTarget = _Form.DataSources.DBDataSources.Item("@LOANMASTERDETAILSC2");
                dsTarget.Clear();
                int row = 0;
                DateTime? previousBookedMonth = null;
                double monthlyInterestTotal = 0;

                while (!rsInterest.EoF)
                {
                    DateTime slabFrom = Convert.ToDateTime(rsInterest.Fields.Item("U_FromDate").Value);
                    DateTime slabTo = Convert.ToDateTime(rsInterest.Fields.Item("U_ToDate").Value);
                    double rate = Convert.ToDouble(rsInterest.Fields.Item("U_InterestPer").Value);
                    double defaultBal = Convert.ToDouble(rsInterest.Fields.Item("U_BalanceAmount").Value);

                    DateTime currentStart = slabFrom;

                    // --- INNER LOOP TO SPLIT SLABS BY MONTH OR REPAYMENT ---
                    while (currentStart <= slabTo)
                    {
                        // Find end of current month
                        DateTime endOfMonth = new DateTime(currentStart.Year, currentStart.Month, DateTime.DaysInMonth(currentStart.Year, currentStart.Month));

                        // Find if there's a repayment within this specific segment
                        var nextRepay = repayments.Where(x => x.PaymentDate > currentStart && x.PaymentDate <= slabTo).OrderBy(x => x.PaymentDate).FirstOrDefault();

                        // Determine segment end (Whichever comes first: Slab end, Month end, or Repayment day-1)
                        DateTime segmentEnd = slabTo;
                        if (endOfMonth < segmentEnd) segmentEnd = endOfMonth;
                        if (nextRepay != null && nextRepay.PaymentDate.AddDays(-1) < segmentEnd) segmentEnd = nextRepay.PaymentDate.AddDays(-1);

                        // Calculate Principal for this segment
                        double principal = GetPrincipalForDate(repayments, currentStart, defaultBal);
                        DateTime interestBookedDate = new DateTime(segmentEnd.Year, segmentEnd.Month, DateTime.DaysInMonth(segmentEnd.Year, segmentEnd.Month));

                        // If month changes, insert the monthly summary row
                        if (previousBookedMonth.HasValue && previousBookedMonth.Value != interestBookedDate)
                        {
                            //AddSummaryRow(dsTarget, ref row, previousBookedMonth.Value, ref monthlyInterestTotal);
                        }
                        previousBookedMonth = interestBookedDate;

                        // Check if this segment was already booked in SAP
                        var existingJE = bookedEntries.FirstOrDefault(e => e.FromDate == currentStart && e.ToDate == segmentEnd);

                        int days = (segmentEnd - currentStart).Days + 1;
                        double interestAmt = (existingJE != null) ? existingJE.InterestAmt : (principal * rate / 100.0) * (days / 365.0);

                        dsTarget.InsertRecord(row);
                        dsTarget.SetValue("U_FromDate", row, currentStart.ToString("yyyyMMdd"));
                        dsTarget.SetValue("U_ToDate", row, segmentEnd.ToString("yyyyMMdd"));
                        dsTarget.SetValue("U_InterestPer", row, rate.ToString());
                        dsTarget.SetValue("U_NoDays", row, days.ToString());
                        dsTarget.SetValue("U_PrincipalAmtC", row, principal.ToString());
                        dsTarget.SetValue("U_InterestAmtC", row, interestAmt.ToString());
                        dsTarget.SetValue("U_IntrDateBook", row, interestBookedDate.ToString("yyyyMMdd"));
                        dsTarget.SetValue("U_DIntrDateBook", row, interestBookedDate.ToString("yyyyMMdd"));


                        dsTarget.SetValue("U_NFromDate", row, GetNepaliDate(currentStart));
                        dsTarget.SetValue("U_NToDate", row, GetNepaliDate(segmentEnd));
                        dsTarget.SetValue("U_NIntrDateBook", row, GetNepaliDate(interestBookedDate));
                        if (existingJE != null) dsTarget.SetValue("U_JENo", row, existingJE.JENo);

                        monthlyInterestTotal += interestAmt;
                        row++;

                        currentStart = segmentEnd.AddDays(1); // Move to next day
                    }
                    rsInterest.MoveNext();
                }

                if (previousBookedMonth.HasValue)
                    //AddSummaryRow(dsTarget, ref row, previousBookedMonth.Value, ref monthlyInterestTotal);

                mtx2 = (SAPbouiCOM.Matrix)_Form.Items.Item("mtxProv").Specific;
                mtx2.LoadFromDataSource();
                //_Form.Mode = BoFormMode.fm_UPDATE_MODE;
                //_Form.Items.Item("1").Click();
            }
            catch (Exception ex) { Utilities.ShowErrorMessage("Error: " + ex.Message); }
            finally
            {
                if (rsInterest != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsInterest);
                if (rsRepay != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsRepay);
                if (rsExisting != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsExisting);
                _Form.Freeze(false);
            }
        }


        private void CopyFromInterestChange()
        {
            SAPbobsCOM.Recordset rsInterest = null, rsRepay = null, rsExisting = null;

            try
            {
                _Form.Freeze(true);

                rsInterest = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rsRepay = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rsExisting = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string loanCode = _HeaderDS.GetValue("U_LoanCode", 0).Trim();
                if (string.IsNullOrWhiteSpace(loanCode)) return;

                // 1. FETCH ALREADY POSTED ENTRIES
                var bookedEntries = new List<dynamic>();
                string sqlCheckExisting = $@"SELECT ""U_FromDate"", ""U_ToDate"", ""U_InterestAmtC"", ""U_JENo"" 
                    FROM ""@LOANMASTERDETAILSC2"" 
                    WHERE ""DocEntry"" = (SELECT ""DocEntry"" FROM ""@LOANMASTERDETAILSH"" WHERE ""U_LoanCode"" = '{loanCode.Replace("'", "''")}')
                    AND IFNULL(""U_JENo"", '') <> ''";

                rsExisting.DoQuery(sqlCheckExisting);
                while (!rsExisting.EoF)
                {
                    bookedEntries.Add(new
                    {
                        FromDate = Convert.ToDateTime(rsExisting.Fields.Item("U_FromDate").Value),
                        ToDate = Convert.ToDateTime(rsExisting.Fields.Item("U_ToDate").Value),
                        InterestAmt = Convert.ToDouble(rsExisting.Fields.Item("U_InterestAmtC").Value),
                        JENo = rsExisting.Fields.Item("U_JENo").Value.ToString()
                    });
                    rsExisting.MoveNext();
                }

                // 2. FETCH REPAYMENTS (Ensuring latest save from DB)
                string sqlRepay = $@"
            SELECT A.""U_PaymentDate"", A.""U_UnPaidAmountC"" AS ""OpeningPrincipal"", C.""U_BalanceAmountC"" AS ""ClosingPrincipal""
            FROM (SELECT T1.""U_PaymentDate"", T1.""U_UnPaidAmountC"", ROW_NUMBER() OVER (PARTITION BY T1.""U_PaymentDate"" ORDER BY T1.""LineId"" ASC) AS RN
                  FROM ""@LOANMASTERDETAILSH"" T0 INNER JOIN ""@LOANMASTERDETAILSC"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                  WHERE T0.""U_LoanCode"" = '{loanCode.Replace("'", "''")}' AND T1.""U_PaymentDate"" IS NOT NULL) A
            INNER JOIN (SELECT T1.""U_PaymentDate"", T1.""U_BalanceAmountC"", ROW_NUMBER() OVER (PARTITION BY T1.""U_PaymentDate"" ORDER BY T1.""LineId"" DESC) AS RN
                        FROM ""@LOANMASTERDETAILSH"" T0 INNER JOIN ""@LOANMASTERDETAILSC"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                        WHERE T0.""U_LoanCode"" = '{loanCode.Replace("'", "''")}' AND T1.""U_PaymentDate"" IS NOT NULL) C ON A.""U_PaymentDate"" = C.""U_PaymentDate""
            WHERE A.RN = 1 AND C.RN = 1 ORDER BY A.""U_PaymentDate"" ASC";
                rsRepay.DoQuery(sqlRepay);

                var repayments = new List<dynamic>();
                while (!rsRepay.EoF)
                {
                    repayments.Add(new
                    {
                        PaymentDate = Convert.ToDateTime(rsRepay.Fields.Item("U_PaymentDate").Value),
                        OpeningPrincipal = Convert.ToDouble(rsRepay.Fields.Item("OpeningPrincipal").Value),
                        ClosingPrincipal = Convert.ToDouble(rsRepay.Fields.Item("ClosingPrincipal").Value)
                    });
                    rsRepay.MoveNext();
                }

                // 3. FETCH INTEREST RATE SLABS
                string sqlInterest = $@"SELECT T1.""U_FromDate"", T1.""U_ToDate"", T1.""U_InterestPer"", T0.""U_BalanceAmountC""
                FROM ""@LOANMASTERDETAILSH"" T0
                INNER JOIN ""@LOANMASTERDETAILSC1"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                WHERE T0.""U_LoanCode"" = '{loanCode.Replace("'", "''")}' AND T1.""U_FromDate"" IS NOT NULL
                ORDER BY T1.""U_FromDate"" ASC";
                rsInterest.DoQuery(sqlInterest);

                SAPbouiCOM.DBDataSource dsTarget = _Form.DataSources.DBDataSources.Item("@LOANMASTERDETAILSC2");
                dsTarget.Clear();
                int row = 0;
                DateTime? previousBookedMonthEnd = null;
                double monthlyInterestTotal = 0;

                while (!rsInterest.EoF)
                {
                    DateTime slabFrom = Convert.ToDateTime(rsInterest.Fields.Item("U_FromDate").Value);
                    DateTime slabTo = Convert.ToDateTime(rsInterest.Fields.Item("U_ToDate").Value);
                    double rate = Convert.ToDouble(rsInterest.Fields.Item("U_InterestPer").Value);
                    double defaultBal = Convert.ToDouble(rsInterest.Fields.Item("U_BalanceAmountC").Value);

                    DateTime currentStart = slabFrom;

                    while (currentStart <= slabTo)
                    {
                        DateTime endOfNepaliMonth = GetNepaliMonthEnd(currentStart);

                        // Find next repayment AFTER current start
                        var nextRepay = repayments.Where(x => x.PaymentDate > currentStart && x.PaymentDate <= slabTo).OrderBy(x => x.PaymentDate).FirstOrDefault();

                        // Determine segment end and Remark
                        DateTime segmentEnd = slabTo;
                        string remark = "Due to Interest Period Change";

                        if (endOfNepaliMonth < segmentEnd)
                        {
                            segmentEnd = endOfNepaliMonth;
                            remark = "End of Nepali Month";
                        }

                        //if (nextRepay != null && nextRepay.PaymentDate.AddDays(-1) < segmentEnd)
                        //{
                        //    segmentEnd = nextRepay.PaymentDate.AddDays(-1);
                        //    remark = "Interest before Repayment on " + nextRepay.PaymentDate.ToString("dd/MM/yyyy");
                        //}

                        if (nextRepay != null && nextRepay.PaymentDate < segmentEnd)
                        {
                            segmentEnd = nextRepay.PaymentDate;
                            remark = "Interest before Repayment on " + nextRepay.PaymentDate.ToString("dd/MM/yyyy");
                        }

                        double principal = GetPrincipalForDate(repayments, currentStart, defaultBal);
                        DateTime interestBookedDate = endOfNepaliMonth;

                        // Summary Row Logic
                        if (previousBookedMonthEnd.HasValue && previousBookedMonthEnd.Value != interestBookedDate)
                        {
                            string nepMonthLabel = GetNepaliDate(previousBookedMonthEnd.Value).Substring(0, 7);
                            AddSummaryRow(dsTarget, ref row, nepMonthLabel, ref monthlyInterestTotal, previousBookedMonthEnd.Value);
                        }
                        previousBookedMonthEnd = interestBookedDate;

                        // Check for existing JE
                        var existingJE = bookedEntries.FirstOrDefault(e => e.FromDate == currentStart && e.ToDate == segmentEnd);

                        int days = (segmentEnd - currentStart).Days + 1;
                        int yearDays = GetTotalDaysInYear(currentStart);
                        double interestAmt = (existingJE != null) ? existingJE.InterestAmt : (principal * rate / 100.0) * (days / 365.0);
                        //in case of leap year
                        //double interestAmt = (existingJE != null) ? existingJE.InterestAmt : (principal * rate / 100.0) * (days / (double)yearDays);

                        dsTarget.InsertRecord(row);
                        dsTarget.SetValue("U_FromDate", row, currentStart.ToString("yyyyMMdd"));
                        dsTarget.SetValue("U_ToDate", row, segmentEnd.ToString("yyyyMMdd"));
                        dsTarget.SetValue("U_InterestPer", row, rate.ToString());
                        dsTarget.SetValue("U_NoDays", row, days.ToString());
                        dsTarget.SetValue("U_PrincipalAmtC", row, principal.ToString());
                        dsTarget.SetValue("U_InterestAmtC", row, interestAmt.ToString());
                        dsTarget.SetValue("U_IntrDateBook", row, interestBookedDate.ToString("yyyyMMdd"));
                        dsTarget.SetValue("U_DIntrDateBook", row, interestBookedDate.ToString("yyyyMMdd"));
                        dsTarget.SetValue("U_Remarks", row, remark); // <--- ADDING REMARK HERE

                        dsTarget.SetValue("U_NFromDate", row, GetNepaliDate(currentStart));
                        dsTarget.SetValue("U_NToDate", row, GetNepaliDate(segmentEnd));
                        dsTarget.SetValue("U_NIntrDateBook", row, GetNepaliDate(interestBookedDate));

                        if (existingJE != null) dsTarget.SetValue("U_JENo", row, existingJE.JENo);

                        monthlyInterestTotal += interestAmt;
                        row++;

                        currentStart = segmentEnd.AddDays(1);
                    }
                    rsInterest.MoveNext();
                }

                if (previousBookedMonthEnd.HasValue)
                {
                    string lastNepMonth = GetNepaliDate(previousBookedMonthEnd.Value).Substring(0, 7);
                    AddSummaryRow(dsTarget, ref row, lastNepMonth, ref monthlyInterestTotal, previousBookedMonthEnd.Value);
                }

                mtx2.LoadFromDataSource();
                AddSequenceNumbersToMatrix(mtx2, dsTarget);
                Form.Mode = BoFormMode.fm_UPDATE_MODE;
                _Form.Items.Item("1").Click();
            }
            catch (Exception ex) { Utilities.ShowErrorMessage("Error: " + ex.Message); }
            finally
            {
                if (rsInterest != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsInterest);
                if (rsRepay != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsRepay);
                if (rsExisting != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsExisting);
                _Form.Freeze(false);
            }
        }

        #endregion

        #region Get Correct Principal for Date
        private double GetPrincipalForDate(List<dynamic> repayments, DateTime date, double defaultBalance)
        {
            // Find the latest repayment that happened ON or BEFORE this date
            var lastRepay = repayments.Where(x => x.PaymentDate < date).OrderByDescending(x => x.PaymentDate).FirstOrDefault();
            if (lastRepay != null) return lastRepay.ClosingPrincipal;

            // If no repayment happened yet, check if there's an "Opening Principal" in the first repayment record
            var firstRepay = repayments.OrderBy(x => x.PaymentDate).FirstOrDefault();
            if (firstRepay != null && date < firstRepay.PaymentDate) return firstRepay.OpeningPrincipal;

            return defaultBalance;
        }
        #endregion

        private Dictionary<int, int> _yearDaysCache = new Dictionary<int, int>();

        #region Get Total Days in Year from Posting Periods (OFPR) (LEAP YEAR CONCEPT)
        private int GetTotalDaysInYear(DateTime date)
        {
            int year = date.Year;
            if (_yearDaysCache.ContainsKey(year)) return _yearDaysCache[year];

            int totalDays = 365; // fallback if no periods are defined for that year
            SAPbobsCOM.Recordset rs = null;
            try
            {
                rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string yStart = $"{year}-01-01";
                string yEnd = $"{year}-12-31";

                string sql = $@"
            SELECT SUM(
                DAYS_BETWEEN(
                    CASE WHEN ""F_RefDate"" < TO_DATE('{yStart}') THEN TO_DATE('{yStart}') ELSE ""F_RefDate"" END,
                    CASE WHEN ""T_RefDate"" > TO_DATE('{yEnd}')   THEN TO_DATE('{yEnd}')   ELSE ""T_RefDate"" END
                ) + 1
            ) AS ""TotalDays""
            FROM OFPR
            WHERE ""F_RefDate"" <= TO_DATE('{yEnd}')
              AND ""T_RefDate"" >= TO_DATE('{yStart}')";

                rs.DoQuery(sql);

                if (!rs.EoF)
                {
                    object val = rs.Fields.Item("TotalDays").Value;
                    int result = Convert.ToInt32(val ?? 0);
                    if (result > 0) totalDays = result;
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage("Error reading posting periods: " + ex.Message);
            }
            finally
            {
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }

            _yearDaysCache[year] = totalDays;
            return totalDays;
        }
        #endregion


        #region Add Sequence Number to Matrix
        private void AddSequenceNumbersToMatrix(SAPbouiCOM.Matrix oMatrix, SAPbouiCOM.DBDataSource oDBDS)
        {
            try
            {
                for (int i = 0; i < oDBDS.Size; i++)
                {
                    // Update the Serial Number column in the DataSource
                    oDBDS.SetValue("U_SrNo", i, (i + 1).ToString());
                }
                oMatrix.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                // Handle error or log
            }
        }
        #endregion

        #region Update Paid and Balance Amount 
        private void UpdatePaidAndBalanceAmounts()
        {
            SAPbobsCOM.Recordset rs = null, rs2 = null;
            try
            {
                string loanCode = _HeaderDS.GetValue("U_LoanCode", 0).Trim();
                rs2 = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string sql1 = $@"SELECT T0.""U_GLAccount"" from ""@LOANMASTERC"" T0 WHERE T0.""U_LoanCode"" = '{loanCode}' ";
                rs2.DoQuery(sql1);
                string glAcc = rs2.Fields.Item("U_GLAccount").Value;

                double totalLoan = 0;
                double.TryParse(_HeaderDS.GetValue("U_LoanAmountC", 0).Trim(), out totalLoan);

                if (string.IsNullOrEmpty(loanCode)) return;

                rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string sql = $@"SELECT SUM(T0.""Debit"") as ""TotalPaid"" FROM JDT1 T0 
                        WHERE (T0.""ProfitCode"" = '{loanCode.Replace("'", "''")}' 
                           OR T0.""OcrCode2"" = '{loanCode.Replace("'", "''")}' 
                           OR T0.""OcrCode3"" = '{loanCode.Replace("'", "''")}' 
                           OR T0.""OcrCode4"" = '{loanCode.Replace("'", "''")}' 
                           OR T0.""OcrCode5"" = '{loanCode.Replace("'", "''")}') 
                        AND T0.""Account"" = '{glAcc}'";

                rs.DoQuery(sql);

                double paidAmount = 0;
                if (!rs.EoF)
                {
                    paidAmount = Convert.ToDouble(rs.Fields.Item("TotalPaid").Value);
                }

                // 3. Calculate Balance
                double balanceAmount = totalLoan - paidAmount;

                // 4. Update Header DataSource
                _HeaderDS.SetValue("U_PaidAmountC", 0, paidAmount.ToString());
                _HeaderDS.SetValue("U_BalanceAmountC", 0, balanceAmount.ToString());

            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage("Error calculating paid amount: " + ex.Message);
            }
            finally
            {
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }

        #endregion

        #region Update RePayment tab details and Paid and Balance Amount
        private void UpdateLoanRepaymentMatrix()
        {
            SAPbobsCOM.Recordset rs = null, rsCheck = null;
            try
            {
                _Form.Freeze(true);

                string loanCode = _HeaderDS.GetValue("U_LoanCode", 0).Trim();
                string docEntry = _HeaderDS.GetValue("DocEntry", 0).Trim();


                double originalBalAmount = 0;
                double.TryParse(_HeaderDS.GetValue("U_BalanceAmountC", 0).Trim(), out originalBalAmount);

                if (string.IsNullOrEmpty(loanCode)) return;

                HashSet<string> existingJEs = new HashSet<string>();
                double totalPaid = 0;
                double runningBalance = originalBalAmount; // This will decrease as we find payments

                // Scan ALREADY SAVED records in the database for this document
                rsCheck = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string sqlExisting = $@"SELECT ""U_JENO"", ""U_BalanceAmountC"", ""U_PaymentAmount"" FROM ""@LOANMASTERDETAILSC"" 
                                WHERE ""DocEntry"" = '{docEntry}' ORDER BY ""U_SrNo"" ASC";
                rsCheck.DoQuery(sqlExisting);

                while (!rsCheck.EoF)
                {
                    string je = rsCheck.Fields.Item("U_JENO").Value.ToString();
                    if (!string.IsNullOrEmpty(je))
                    {
                        existingJEs.Add(je);

                        double savedPayment = Convert.ToDouble(rsCheck.Fields.Item("U_PaymentAmount").Value);
                        totalPaid += savedPayment;
                        // Update runningBalance based on what is already saved in history
                        runningBalance = Convert.ToDouble(rsCheck.Fields.Item("U_BalanceAmountC").Value);
                    }
                    rsCheck.MoveNext();
                }

                // Get GL Account
                rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string sqlGLAcc = $@"SELECT T0.""U_GLAccount"" from ""@LOANMASTERC"" T0 WHERE T0.""U_LoanCode"" = '{loanCode.Replace("'", "''")}' ";
                rs.DoQuery(sqlGLAcc);
                if (rs.EoF) return;
                string glAcc = rs.Fields.Item("U_GLAccount").Value.ToString();

                // Fetch all payments from GL
                //string sqlRepay = $@"SELECT SUM(T0.""Debit"") as ""Total"", T1.""RefDate"", T1.""TransId""
                //             FROM JDT1 T0 
                //             INNER JOIN OJDT T1 ON T0.""TransId"" = T1.""TransId""
                //             WHERE (T0.""ProfitCode"" = '{loanCode.Replace("'", "''")}' 
                //                OR T0.""OcrCode2"" = '{loanCode.Replace("'", "''")}' 
                //                OR T0.""OcrCode3"" = '{loanCode.Replace("'", "''")}' 
                //                OR T0.""OcrCode4"" = '{loanCode.Replace("'", "''")}' 
                //                OR T0.""OcrCode5"" = '{loanCode.Replace("'", "''")}') 
                //             AND T0.""Account"" = '{glAcc}' and T0.""Debit"">0
                //             GROUP BY T1.""RefDate"", T1.""TransId""
                //             ORDER BY T1.""RefDate"" ASC";
                //rs.DoQuery(sqlRepay);
                string sqlRepay = $@"
                SELECT 
                    SUM(T0.""Debit"") AS ""Total"", T1.""RefDate"", T1.""TransId""
                FROM JDT1 T0
                INNER JOIN OJDT T1 ON T0.""TransId"" = T1.""TransId""
                WHERE 
                    (T0.""ProfitCode"" = '{loanCode.Replace("'", "''")}' 
                     OR T0.""OcrCode2"" = '{loanCode.Replace("'", "''")}' 
                     OR T0.""OcrCode3"" = '{loanCode.Replace("'", "''")}' 
                     OR T0.""OcrCode4"" = '{loanCode.Replace("'", "''")}' 
                     OR T0.""OcrCode5"" = '{loanCode.Replace("'", "''")}')
        
                    AND T0.""Account"" = '{glAcc}' AND T0.""Debit"" > 0 AND IFNULL(T1.""StornoToTr"", 0) = 0 
                    AND T1.""TransId"" NOT IN (
                        SELECT DISTINCT A.""StornoToTr""
                        FROM OJDT A
                        WHERE IFNULL(A.""StornoToTr"", 0) <> 0
                    )

                GROUP BY T1.""RefDate"", T1.""TransId""
                ORDER BY T1.""RefDate"" ASC";
                rs.DoQuery(sqlRepay);

                bool newRowsAdded = false;

                // Append only NEW entries
                while (!rs.EoF)
                {
                    string currentJE = rs.Fields.Item("TransId").Value.ToString();

                    if (!existingJEs.Contains(currentJE))
                    {
                        newRowsAdded = true;

                        int newRowIdx = _ChildDS2.Size;
                        if (newRowIdx == 1 && string.IsNullOrEmpty(_ChildDS2.GetValue("U_JENO", 0)))
                            newRowIdx = 0;
                        else
                            _ChildDS2.InsertRecord(newRowIdx);

                        double newPaymentAmount = Convert.ToDouble(rs.Fields.Item("Total").Value);

                        // Logic: Unpaid = current running balance (from previous row)
                        double unpaidBefore = runningBalance;

                        // Subtract current payment to get new balance
                        runningBalance = unpaidBefore - newPaymentAmount;

                        // Add to the total paid sum for the header
                        totalPaid += newPaymentAmount;

                        DateTime payDate = (DateTime)rs.Fields.Item("RefDate").Value;

                        _ChildDS2.SetValue("U_SrNo", newRowIdx, (newRowIdx + 1).ToString());
                        _ChildDS2.SetValue("U_PaymentDate", newRowIdx, ((DateTime)rs.Fields.Item("RefDate").Value).ToString("yyyyMMdd"));
                        _ChildDS2.SetValue("U_JENO", newRowIdx, currentJE);
                        _ChildDS2.SetValue("U_UnpaidAmountC", newRowIdx, unpaidBefore.ToString());
                        _ChildDS2.SetValue("U_PaymentAmountC", newRowIdx, newPaymentAmount.ToString());
                        _ChildDS2.SetValue("U_BalanceAmountC", newRowIdx, runningBalance.ToString());
                        _ChildDS2.SetValue("U_NPaymentDate", newRowIdx, GetNepaliDate(payDate));
                    }
                    rs.MoveNext();
                }

                // Final Header Synchronization
                if (newRowsAdded)
                {
                    mtx3.LoadFromDataSource();

                    // Push the final calculated balance to the header
                    _HeaderDS.SetValue("U_PaidAmountC", 0, totalPaid.ToString());
                    _HeaderDS.SetValue("U_BalanceAmountC", 0, runningBalance.ToString());

                    if (_Form.Mode == BoFormMode.fm_OK_MODE)
                        _Form.Mode = BoFormMode.fm_UPDATE_MODE;
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage("Error: " + ex.Message);
            }
            finally
            {
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                if (rsCheck != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsCheck);
                _Form.Freeze(false);
            }
        }
        #endregion

        #region Post Interest Provision JE
        private void PostInterestProvisionJE()
        {
            SAPbobsCOM.JournalEntries oJE = null;
            SAPbobsCOM.Recordset rs = null;

            try
            {
                _Form.Freeze(true);
                string loanCode = _HeaderDS.GetValue("U_LoanCode", 0).Trim();
                if (string.IsNullOrEmpty(loanCode)) return;

                // 1. Fetch Accounts and Dimension info once
                rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                //    string sqlQuery = $@"
                //SELECT 
                //    T0.""U_ProvisionAccount"", 
                //    T0.""U_InterestAccount"",  
                //    T1.""DimCode"" 
                //FROM ""@LOANMASTERC"" T0
                //LEFT JOIN OOCR T1 ON T0.""U_LoanCode"" = T1.""OcrCode""
                //WHERE T0.""U_LoanCode"" = '{loanCode.Replace("'", "''")}'";

                string sqlQuery = $@"
                    SELECT 
                        T0.""U_ProvisionAc"" ""Provision"", 
                         CASE WHEN T1.""U_LoanType"" = 'TR Loan' THEN T0.""U_TRInterestAc"" WHEN T1.""U_LoanType"" = 'Term Loan' THEN T0.""U_TermInterestAc""
                    WHEN T1.""U_LoanType"" = 'Bridge Gap Loan' THEN T0.""U_BridgeInterestAc"" WHEN T1.""U_LoanType"" = 'Demand Loan' THEN T0.""U_DemandInterestAc""
                    WHEN T1.""U_LoanType"" = 'Over Draft Loan' THEN T0.""U_ODInterestAc"" WHEN T1.""U_LoanType"" = 'Short Term Loan' THEN T0.""U_STLInterestAc""
                    END AS ""Interest"", T5.""DimCode"", T2.""U_LoanCode""
                    FROM ""@LOANGLAC"" T0
                    INNER JOIN ""@LOANMASTERH"" T4 ON T0.""Code"" = T4.""U_BankCode""
                    INNER JOIN ""@LOANMASTERC"" T1 ON T1.""DocEntry"" = T4.""DocEntry""
                    INNER JOIN ""@LOANMASTERDETAILSH"" T2 ON T1.""U_LoanCode"" = T2.""U_LoanCode""
                    LEFT JOIN OOCR T5 ON T1.""U_LoanCode"" = T5.""OcrCode""
                    WHERE T1.""U_LoanCode"" = '{loanCode.Replace("'", "''")}'";

                rs.DoQuery(sqlQuery);

                if (rs.EoF)
                {
                    Utilities.ShowErrorMessage("Loan configuration or Distribution Rule not found.");
                    return;
                }

                //interest debit provision cr prviously
                string debitAcc = rs.Fields.Item("Interest").Value.ToString();
                string creditAcc = rs.Fields.Item("Provision").Value.ToString();
                int dimension = Convert.ToInt32(rs.Fields.Item("DimCode").Value);
                string currentMonthKey = DateTime.Today.ToString("yyyyMM");

                string todayNepDate = GetNepaliDate(DateTime.Today);
                string currentNepMonthKey = todayNepDate.Substring(0, 7);

                int successCount = 0;
                bool isDataUpdated = false;

                // Iterate through each row in the matrix datasource
                for (int i = 0; i < _ChildDS1.Size; i++)
                {
                    string toDateStr = _ChildDS1.GetValue("U_ToDate", i).Trim();
                    string intrstBook = _ChildDS1.GetValue("U_IntrDateBook", i).Trim();
                    string interestAmtStr = _ChildDS1.GetValue("U_InterestAmtC", i).Trim();
                    string existingJENo = _ChildDS1.GetValue("U_JENo", i).Trim();
                    string dateBook = _ChildDS1.GetValue("U_IntrDateBook", i).Trim();

                    // Skip if row is incomplete, amount is zero, or already booked
                    if (string.IsNullOrEmpty(dateBook) || string.IsNullOrEmpty(interestAmtStr) || !string.IsNullOrEmpty(existingJENo))
                        continue;

                    double interestAmt = Convert.ToDouble(interestAmtStr);
                    if (interestAmt <= 0) continue;

                    DateTime rowToDate = DateTime.ParseExact(dateBook, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                    string rowMonthKey = rowToDate.ToString("yyyyMM");

                    // --- NEPALI DATE VALIDATION ---
                    // Get Row Nepali Month Key 
                    string rowNepDate = GetNepaliDate(rowToDate);
                    string rowNepMonthKey = rowNepDate.Substring(0, 7);

                    // POST ONLY PREVIOUS Nepali Months
                    // If row month is >= current month, skip it.
                    if (string.Compare(rowNepMonthKey, currentNepMonthKey) >= 0)
                        continue;

                    DateTime rowPostDate = DateTime.ParseExact(intrstBook, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);

                    // Skip future months on the basis of english date 
                    //if (string.Compare(rowMonthKey, currentMonthKey) > 0)
                    //    continue;

                    ////Skip future dates
                    //if (rowToDate > DateTime.Today)
                    //    continue;

                    // Calculate Last Day of the month for this specific row
                    //DateTime postingDate = new DateTime(rowToDate.Year, rowToDate.Month, DateTime.DaysInMonth(rowToDate.Year, rowToDate.Month));
                    DateTime postingDate = DateTime.ParseExact(intrstBook, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);

                    // 3. Create a Journal Entry for this specific row
                    oJE = (SAPbobsCOM.JournalEntries)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                    oJE.ReferenceDate = postingDate;
                    oJE.Memo = $"Interest Provision - {loanCode} - Row {i + 1}";

                    // Add Debit Line (Interest Expense)
                    AddJELine(oJE, debitAcc, interestAmt, 0, 1, loanCode, dimension,false);
                    // Add Credit Line (Interest Provision)
                    AddJELine(oJE, creditAcc, 0, interestAmt, 1, loanCode, dimension,true);

                    int errCode = oJE.Add();

                    if (errCode != 0)
                    {
                        string errMsg = Utilities.Application.Company.GetLastErrorDescription();
                        Utilities.ShowErrorMessage($"Error posting row {i + 1}: {errMsg}");
                        // Break or Continue? Usually, in finance, we break to avoid partial messy postings.
                        break;
                    }
                    else
                    {
                        string newJeKey = Utilities.Application.Company.GetNewObjectKey();
                        _ChildDS1.SetValue("U_JENo", i, newJeKey);

                        try
                        {
                            // Get the dates of the segment currently being booked
                            string segmentFrom = _ChildDS1.GetValue("U_FromDate", i).Trim();
                            string segmentTo = _ChildDS1.GetValue("U_ToDate", i).Trim();

                            if (!string.IsNullOrEmpty(segmentFrom) && !string.IsNullOrEmpty(segmentTo))
                            {
                                DateTime segFrm = DateTime.ParseExact(segmentFrom, "yyyyMMdd", null);
                                DateTime segTo = DateTime.ParseExact(segmentTo, "yyyyMMdd", null);

                                // Loop through the Slab DataSource (C1) to find which slab this segment belongs to
                                for (int k = 0; k < _ChildDS.Size; k++)
                                {
                                    string slabFromStr = _ChildDS.GetValue("U_FromDate", k).Trim();
                                    string slabToStr = _ChildDS.GetValue("U_ToDate", k).Trim();

                                    if (!string.IsNullOrEmpty(slabFromStr) && !string.IsNullOrEmpty(slabToStr))
                                    {
                                        DateTime slabFrm = DateTime.ParseExact(slabFromStr, "yyyyMMdd", null);
                                        DateTime slabTo = DateTime.ParseExact(slabToStr, "yyyyMMdd", null);

                                        // If the booked segment falls within this Slab's date range
                                        if (segFrm >= slabFrm && segTo <= slabTo)
                                        {
                                            _ChildDS.SetValue("U_Status", k, "C");
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            // Log error but don't stop the JE posting process
                        }

                        successCount++;
                        isDataUpdated = true;
                    }
                    if (oJE != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oJE);
                    oJE = null;
                }

                //Update UI if JEs were created
                if (isDataUpdated)
                {
                   _Matrix.LoadFromDataSource();
                    mtx2.LoadFromDataSource();
                    if (_Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        _Form.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;

                    _Form.Items.Item("1").Click(); // Save the JE Numbers to the UDT

                    RefreshFormViaFind(loanCode);

                    Utilities.Application.SBO_Application.SetStatusBarMessage($"Successfully posted {successCount} individual Journal Entries.", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage("PostInterestProvisionJE Error: " + ex.Message);
            }
            finally
            {
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                if (oJE != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oJE);
                _Form.Freeze(false);
            }
        }
        #endregion

        private void AddJELine(SAPbobsCOM.JournalEntries oJE, string accCode, double debit, double credit, int branchId, string distRule, int dim, bool tagCostCenter)
        {
            oJE.Lines.AccountCode = accCode;
            oJE.Lines.Debit = debit;
            oJE.Lines.Credit = credit;
            oJE.Lines.BPLID = branchId;

            // Only apply cost center logic if tagCostCenter is true
            if (tagCostCenter)
            {
                switch (dim)
                {
                    case 1: oJE.Lines.CostingCode = distRule; break;
                    case 2: oJE.Lines.CostingCode2 = distRule; break;
                    case 3: oJE.Lines.CostingCode3 = distRule; break;
                    case 4: oJE.Lines.CostingCode4 = distRule; break;
                    case 5: oJE.Lines.CostingCode5 = distRule; break;
                }
            }

            oJE.Lines.Add();
        }

        private void AddSummaryRow(SAPbouiCOM.DBDataSource ds, ref int row, string nepaliMonthLabel, ref double total, DateTime interestBookedDate)
        {
            string monthName = GetNepaliMonthName(nepaliMonthLabel);

            ds.InsertRecord(row);
            ds.SetValue("U_FromDate", row, "");
            ds.SetValue("U_ToDate", row, "");
            ds.SetValue("U_DIntrDateBook", row, interestBookedDate.ToString("yyyyMMdd"));
            // Update the label to reflect Nepali month
            string description = $"Monthly Total Interest for (BS) {nepaliMonthLabel} ({monthName})";
            //ds.SetValue("U_InterestPer", row, "Monthly Total Interest for (BS) " + nepaliMonthLabel);
            ds.SetValue("U_InterestPer", row, description);
            ds.SetValue("U_NoDays", row, "");
            ds.SetValue("U_PrincipalAmtC", row, "");
            ds.SetValue("U_InterestAmtC", row, "");
            ds.SetValue("U_Remarks", row, total.ToString("N2"));
            row++;
            total = 0;
        }

        private double SafeDouble(object val)
        {
            if (val == null) return 0;
            string s = val.ToString().Trim();
            if (string.IsNullOrEmpty(s)) return 0;

            double result;
            if (double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out result))
                return result;

            double.TryParse(s, out result);
            return result;
        }


        private void LockBookedRows()
        {
            try
            {
                // Get the Max Booked Date
                string docEntry = _HeaderDS.GetValue("DocEntry", 0).Trim();
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery($@"SELECT MAX(T0.""U_ToDate"") FROM ""@LOANMASTERDETAILSC2"" T0 WHERE T0.""DocEntry"" = '{docEntry}' AND IFNULL(T0.""U_JENo"",'') <> ''");

                if (rs.EoF || rs.Fields.Item(0).Value == null || rs.Fields.Item(0).Value.ToString() == "") return;
                DateTime maxBookedDate = (DateTime)rs.Fields.Item(0).Value;

                // Lock Interest Slab Matrix (mtxInt)
                // If the slab's ToDate is <= maxBookedDate, the whole row is locked.
                for (int i = 1; i <= _Matrix.VisualRowCount; i++)
                {
                    string rowToDateStr = _ChildDS.GetValue("U_ToDate", i - 1).Trim();
                    if (string.IsNullOrEmpty(rowToDateStr)) continue;

                    DateTime rowToDate = DateTime.ParseExact(rowToDateStr, "yyyyMMdd", null);
                    if (rowToDate <= maxBookedDate)
                    {
                        _Matrix.CommonSetting.SetRowEditable(i, false);
                    }
                    //else
                    //{
                    //    _Matrix.CommonSetting.SetRowEditable(i, true);
                    //}
                }
            }
            catch { }
        }


        #region Calculate Row Interest
        private void CalculateRowInterest(int rowIdx)
        {
            try
            {
                _Form.Freeze(true);

                // Get cell objects for the specific row
                SAPbouiCOM.EditText oEditBase = (SAPbouiCOM.EditText)_Matrix.Columns.Item("ColBInrst").Cells.Item(rowIdx).Specific;
                SAPbouiCOM.EditText oEditPrem = (SAPbouiCOM.EditText)_Matrix.Columns.Item("ColPInrst").Cells.Item(rowIdx).Specific;
                SAPbouiCOM.EditText oEditTotal = (SAPbouiCOM.EditText)_Matrix.Columns.Item("ColInrst").Cells.Item(rowIdx).Specific;

                double baseVal = 0;
                double premVal = 0;

                double.TryParse(oEditBase.Value, out baseVal);
                double.TryParse(oEditPrem.Value, out premVal);

                double totalVal = baseVal + premVal;

                // Push back to the Total Interest column in the UI
                oEditTotal.Value = totalVal.ToString();

                // Flush to DataSource to ensure the value is saved in @LOANMASTERDETAILSC1
                _Matrix.FlushToDataSource();
            }
            catch (Exception ex)
            {
                // Silent fail or log error
            }
            finally
            {
                _Form.Freeze(false);
            }
        }
        #endregion

        #region Get Nepali Date
        private string GetNepaliDate(DateTime engDate)
        {
            string key = engDate.ToString("yyyyMMdd");

            // Look up the date in the memory dictionary
            if (CalendarCache.ContainsKey(key))
            {
                return CalendarCache[key];
            }

            // Fallback if the date is missing in @NEPALICALENDAR
            return engDate.ToString("yyyy/MM/dd");
        }
        #endregion


        #region Update Nepali in matrix cell
        private void UpdateNepaliDateInCell(int row, string colUID)
        {
            if (_isSyncingDate) return;

            try
            {
                _isSyncingDate = true;

                string targetSourceCol = "";
                string targetMatrixCol = "";

                //Get the Specific Cell that was just validated
                SAPbouiCOM.EditText oEditEng = (SAPbouiCOM.EditText)_Matrix.Columns.Item(colUID).Cells.Item(row).Specific;
                string engDateStr = oEditEng.Value.Trim();

                if (colUID == "ColFrmD")
                {
                    targetSourceCol = "U_NFromDate";
                    targetMatrixCol = "ColFrmDN";
                }
                else if (colUID == "ColToD")
                {
                    targetSourceCol = "U_NToDate";
                    targetMatrixCol = "ColToDN";
                }

                if (!string.IsNullOrEmpty(engDateStr))
                {
                    DateTime engDate = DateTime.ParseExact(engDateStr, "yyyyMMdd", null);
                    string nepaliDate = GetNepaliDate(engDate);

                    // Update the Nepali UI Cell directly
                    SAPbouiCOM.EditText oEditNep = (SAPbouiCOM.EditText)_Matrix.Columns.Item(targetMatrixCol).Cells.Item(row).Specific;
                    oEditNep.Value = nepaliDate;

                    // Update the DBDataSource manually for that specific row/column
                    // This ensures data is saved without flushing the whole matrix
                    _ChildDS.SetValue(targetSourceCol, row - 1, nepaliDate);
                }
            }
            catch (Exception) { }
            finally
            {
                _isSyncingDate = false;
            }
        }
        #endregion

        #region Load Calendar Cache from Nepali Calendar
        private void LoadCalendarCache()
        {
            _calendarCache = new Dictionary<string, string>();
            _nepaliMonthEndCache = new Dictionary<string, DateTime>();
            _nepaliMonthNameCache = new Dictionary<string, string>();

            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                // Order by English date to ensure the last record we find for a month is the month-end
                string sql = "SELECT \"U_EnglishDate\", \"U_NepaliDate\", \"U_NepaliMonth\" FROM \"@NEPALICALENDER\" ORDER BY \"U_EnglishDate\" ASC";
                rs.DoQuery(sql);

                while (!rs.EoF)
                {
                    DateTime engDt = (DateTime)rs.Fields.Item("U_EnglishDate").Value;
                    string nepDt = rs.Fields.Item("U_NepaliDate").Value.ToString(); // Expected format: "2080/01/01"
                    string nepMonth = rs.Fields.Item("U_NepaliMonth").Value.ToString();

                    string key = engDt.ToString("yyyyMMdd");
                    if (!_calendarCache.ContainsKey(key)) _calendarCache.Add(key, nepDt);

                    // Extract YYYY/MM (first 7 chars)
                    if (nepDt.Length >= 7)
                    {
                        string yearMonthKey = nepDt.Substring(0, 7);
                        // Overwrite the entry as we loop; the last date processed for "2080/01" will be its end date.
                        _nepaliMonthEndCache[yearMonthKey] = engDt;

                        if (!_nepaliMonthNameCache.ContainsKey(yearMonthKey))
                        {
                            _nepaliMonthNameCache.Add(yearMonthKey, nepMonth);
                        }
                    }
                    rs.MoveNext();
                }
            }
            catch (Exception ex) { Utilities.ShowErrorMessage("Error loading Nepali Calendar: " + ex.Message); }
            finally { System.Runtime.InteropServices.Marshal.ReleaseComObject(rs); }
        }
        #endregion

        #region Dictionary for Calendar Cache
        private Dictionary<string, string> CalendarCache
        {
            get
            {
                if (_calendarCache == null) LoadCalendarCache();
                return _calendarCache;
            }
        }
        #endregion


        #region Get Nepali Month End
        private DateTime GetNepaliMonthEnd(DateTime engDate)
        {
            string nepDateStr = GetNepaliDate(engDate); // returns "YYYY/MM/DD"
            string yearMonth = nepDateStr.Substring(0, 7); // extracts "YYYY/MM"

            if (_nepaliMonthEndCache != null && _nepaliMonthEndCache.ContainsKey(yearMonth))
            {
                return _nepaliMonthEndCache[yearMonth];
            }
            // Fallback to English month end only if calendar data is missing
            return new DateTime(engDate.Year, engDate.Month, DateTime.DaysInMonth(engDate.Year, engDate.Month));
        }
        #endregion

        #region Get Nepali Month Name
        private string GetNepaliMonthName(string yearMonthKey)
        {
            if (_nepaliMonthNameCache != null && _nepaliMonthNameCache.ContainsKey(yearMonthKey))
            {
                return _nepaliMonthNameCache[yearMonthKey];
            }
            return "";
        }
        #endregion

        #region Validate Interest Slabs
        private bool ValidateInterestSlabs()
        {
            try
            {
                // Force the UI values into the DBDataSource
                _Matrix.FlushToDataSource();

                // In ADD mode, sometimes Size is 0 even if a row is visible.
                // We ensure we at least check what is visible in the Matrix.
                int rowCount = _ChildDS.Size;
                if (rowCount == 0 && _Matrix.VisualRowCount > 0)
                {
                    // If DataSource is out of sync, we try to flush again or 
                    // use VisualRowCount to drive the loop
                    rowCount = _Matrix.VisualRowCount;
                }

                if (rowCount == 0)
                {
                    Utilities.ShowErrorMessage("At least one Interest Slab row is required.");
                    return false;
                }

                for (int i = 0; i < rowCount; i++)
                {
                    // Use GetValue but also handle potential out-of-range if Size != VisualRowCount
                    string fDate = "";
                    string tDate = "";
                    string baseInt = "";
                    string premInt = "";

                    try
                    {
                        fDate = _ChildDS.GetValue("U_FromDate", i).Trim();
                        tDate = _ChildDS.GetValue("U_ToDate", i).Trim();
                        baseInt = _ChildDS.GetValue("U_BaseInterestPer", i).Trim();
                        premInt = _ChildDS.GetValue("U_PremInterestPer", i).Trim();
                    }
                    catch
                    {
                        // If the DataSource is smaller than the loop index, it's an empty ghost row
                        continue;
                    }

   
                    bool isRowEmpty = string.IsNullOrWhiteSpace(fDate) &&
                                      string.IsNullOrWhiteSpace(tDate) &&
                                      string.IsNullOrWhiteSpace(baseInt) &&
                                      string.IsNullOrWhiteSpace(premInt);

                    bool isRowPartiallyFilled = !string.IsNullOrWhiteSpace(fDate) ||
                                                !string.IsNullOrWhiteSpace(tDate) ||
                                                !string.IsNullOrWhiteSpace(baseInt) ||
                                                !string.IsNullOrWhiteSpace(premInt);

                    //FIRST ROW MUST HAVE DATA ---
                    if (i == 0)
                    {
                        if (isRowEmpty)
                        {
                            Utilities.ShowErrorMessage("The first row in Interest Change tab cannot be empty.");
                            _Matrix.Columns.Item("ColFrmD").Cells.Item(1).Click();
                            return false;
                        }
                    }

                    // IF ANY FIELD HAS DATA, ALL MUST HAVE DATA ---
                    if (isRowPartiallyFilled)
                    {
                        if (string.IsNullOrWhiteSpace(fDate))
                        {
                            Utilities.ShowErrorMessage($"Row {i + 1}: From Date is required.");
                            _Matrix.Columns.Item("ColFrmD").Cells.Item(i + 1).Click();
                            return false;
                        }
                        if (string.IsNullOrWhiteSpace(tDate))
                        {
                            Utilities.ShowErrorMessage($"Row {i + 1}: To Date is required.");
                            _Matrix.Columns.Item("ColToD").Cells.Item(i + 1).Click();
                            return false;
                        }
                        if (string.IsNullOrWhiteSpace(baseInt))
                        {
                            Utilities.ShowErrorMessage($"Row {i + 1}: Base Interest % is required.");
                            _Matrix.Columns.Item("ColBInrst").Cells.Item(i + 1).Click();
                            return false;
                        }
                        if (string.IsNullOrWhiteSpace(premInt))
                        {
                            Utilities.ShowErrorMessage($"Row {i + 1}: Premium Interest % is required.");
                            _Matrix.Columns.Item("ColPInrst").Cells.Item(i + 1).Click();
                            return false;
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage("Validation Error: " + ex.Message);
                return false;
            }
        }
        #endregion

        #region Form Refresh after JE creation
        private void RefreshFormViaFind(string loanCode)
        {
            try
            {
                _Form.Freeze(true);
                _Form.Mode = BoFormMode.fm_FIND_MODE;

                // Enable field temporarily to set the find value
                _Form.Items.Item("txtLoanC").Enabled = true;
                ((SAPbouiCOM.EditText)_Form.Items.Item("txtLoanC").Specific).Value = loanCode;

                _Form.Items.Item("1").Click();

                _Form.Items.Item("txtLoanC").Enabled = false;
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage("Refresh Error: " + ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
            }
        }
        #endregion

        private void SortDataSourceChronologically()
        {
            try
            {
                if (_ChildDS1.Size <= 1) return; // Nothing to sort

                // 1. Move data from DataSource to a temporary List
                var rowList = new List<Dictionary<string, string>>();
                for (int i = 0; i < _ChildDS1.Size; i++)
                {
                    var row = new Dictionary<string, string>();
                    // Add all fields that your matrix uses
                    row["U_FromDate"] = _ChildDS1.GetValue("U_FromDate", i).Trim();
                    row["U_ToDate"] = _ChildDS1.GetValue("U_ToDate", i).Trim();
                    row["U_InterestPer"] = _ChildDS1.GetValue("U_InterestPer", i).Trim();
                    row["U_Remarks"] = _ChildDS1.GetValue("U_Remarks", i).Trim();
                    row["U_PrincipalAmtC"] = _ChildDS1.GetValue("U_PrincipalAmtC", i).Trim();
                    row["U_InterestAmtC"] = _ChildDS1.GetValue("U_InterestAmtC", i).Trim();
                    row["U_IntrDateBook"] = _ChildDS1.GetValue("U_IntrDateBook", i).Trim();
                    row["U_DIntrDateBook"] = _ChildDS1.GetValue("U_DIntrDateBook", i).Trim();
                    row["U_JENo"] = _ChildDS1.GetValue("U_JENo", i).Trim();
                    row["U_NFromDate"] = _ChildDS1.GetValue("U_NFromDate", i).Trim();
                    row["U_NToDate"] = _ChildDS1.GetValue("U_NToDate", i).Trim();
                    row["U_NIntrDateBook"] = _ChildDS1.GetValue("U_NIntrDateBook", i).Trim();
                    row["U_NoDays"] = _ChildDS1.GetValue("U_NoDays", i).Trim();

                    rowList.Add(row);
                }

                // 2. Sort the List
                // Primary: Book Date (YYYYMMDD string sort works perfectly)
                // Secondary: Row Type (Detail rows first [FromDate exists], Summary rows last [FromDate empty])
                var sortedList = rowList.OrderBy(r => r["U_DIntrDateBook"])
                                        .ThenBy(r => string.IsNullOrEmpty(r["U_FromDate"]) ? 1 : 0)
                                        .ThenBy(r => r["U_FromDate"])
                                        .ToList();

                // 3. Clear and Refill DataSource
                _ChildDS1.Clear();
                for (int i = 0; i < sortedList.Count; i++)
                {
                    _ChildDS1.InsertRecord(i);
                    foreach (var key in sortedList[i].Keys)
                    {
                        _ChildDS1.SetValue(key, i, sortedList[i][key]);
                    }
                }

                // 4. Load into Matrix
                mtx2.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage("Sort Error: " + ex.Message);
            }
        }
    }
}
