using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BankFacility
{
    class ClsMonthlyInterestCalculation : Base
    {
        #region Constructor and Distructor
        public ClsMonthlyInterestCalculation() { }
        ~ClsMonthlyInterestCalculation() { }
        #endregion

        #region VARIABLE DECLARATION
        public SAPbouiCOM.DBDataSource _MasterDBDataSource = null;
        SAPbouiCOM.ChooseFromListEvent _SysCFLEvent = null;
        #endregion

        public override void FormData_Event(ref BusinessObjectInfo BusinessObjectInfo, ref bool oBubbleEvent)
        {
            if (_Form.Mode == BoFormMode.fm_OK_MODE)
            {
                _Form.Items.Item("txtSDate").Enabled = false;
                _Form.Items.Item("txtBCode").Enabled = false;
                _Form.Items.Item("txtLoanT").Enabled = false;
                _Form.Items.Item("cmbPosP").Enabled = false;
                _Form.Items.Item("cmbFisY").Enabled = false;
            }
        }

        public override void Item_Event(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            #region Before Action
            if (pVal.BeforeAction)
            {
                if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED)
                {
                    if (pVal.ItemUID == "1" || pVal.ItemUID == "btnFD")
                    {
                        if (_Form.Mode == BoFormMode.fm_ADD_MODE)
                        {
                            if (!ValidateFetchCriteria())
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (pVal.ItemUID == "1" && pVal.BeforeAction == true && _Form.Mode == BoFormMode.fm_ADD_MODE)
                            {
                                SAPbouiCOM.Matrix mtx = (SAPbouiCOM.Matrix)_Form.Items.Item("mtxFtch").Specific;
                                mtx.FlushToDataSource();
                                if (mtx.VisualRowCount == 0)
                                {
                                    Utilities.Application.SBO_Application.StatusBar.SetText("Cannot add document: No loan rows found. Please click 'Fetch Data' first.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    BubbleEvent = false;
                                    return;
                                }

                                SAPbouiCOM.DBDataSource dsLines = _Form.DataSources.DBDataSources.Item("@MINTERESTCC");

                                //if (!ValidateSlabGapsLineWise())
                                //{
                                //    BubbleEvent = false;
                                //    return;
                                //}

                                //if (!PostSlabsToLoanMasters())
                                //{
                                //    BubbleEvent = false;
                                //    return;
                                //}
                            }
                        }
                    }
                }

                if (pVal.EventType == BoEventTypes.et_VALIDATE && pVal.ItemUID == "txtSDate")
                {
                    SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)_Form.Items.Item("txtSDate").Specific;
                    string dateStr = oEdit.Value.Trim();

                    if (!string.IsNullOrEmpty(dateStr))
                    {
                        if (!IsFirstDayOfNepaliMonth(dateStr))
                        {
                            Utilities.Application.SBO_Application.StatusBar.SetText("Selection Failed: Start Date must be the 1st day of a Nepali Month.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                            BubbleEvent = false; 
                            return;
                        }
                    }
                }
            }
            #endregion
            else
            {
                switch (pVal.EventType)
                {

                    case BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "btnFD" && pVal.ActionSuccess)
                        {
                            FetchData(false);
                        }
                        //if (pVal.ItemUID == "1" && (_Form.Mode == BoFormMode.fm_ADD_MODE || _Form.Mode == BoFormMode.fm_UPDATE_MODE))
                        //{
                        //    // Add successful, now sync Rate Slabs (C1) only
                        //    PostSlabsToLoanMasters();
                        //}

                        if (pVal.ItemUID == "btnCalc" && pVal.ActionSuccess)
                        {
                            CalculateInterestForAllRows();
                        }
                        if (pVal.ItemUID == "btnIntrst" && pVal.ActionSuccess) 
                        {
                            BookInterestBatch();
                        }
                        break;

                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        _SysCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                        if (pVal.ItemUID == "txtBCode" && _SysCFLEvent.SelectedObjects != null)
                        {
                            try
                            {
                                string bankCode = _SysCFLEvent.SelectedObjects.GetValue("Code", 0).ToString().Trim();
                                _MasterDBDataSource.SetValue("U_BankCode", 0, bankCode);
                                //((SAPbouiCOM.EditText)_Form.Items.Item("txtBCode").Specific).Value = bankCode;

                                //CheckAndLoadExistingRecord();
                                //BoFormMode currentMode = _Form.Mode;
                            }
                            catch
                            {

                            }
                        }
                        if (pVal.ItemUID == "txtLoanT" && _SysCFLEvent.SelectedObjects != null)
                        {
                            try
                            {
                                string loanType = _SysCFLEvent.SelectedObjects.GetValue("Name", 0).ToString().Trim();
                                ((SAPbouiCOM.EditText)_Form.Items.Item("txtLoanT").Specific).Value = loanType;
                            }
                            catch
                            {

                            }
                        }
                        break;

                    case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                        if (pVal.ItemUID == "txtSDate" && pVal.ActionSuccess)
                        {
                            // Auto-calculate the End Date based on the Start Date
                            string startDate = _MasterDBDataSource.GetValue("U_StartDate", 0).Trim();
                            if (!string.IsNullOrEmpty(startDate))
                            {
                                string calculatedEndDate = GetNepaliMonthEndDate(startDate);
                                _MasterDBDataSource.SetValue("U_EndDate", 0, calculatedEndDate);
                                // No need to refresh UI if bound to DBDataSource, but good practice:
                                //((SAPbouiCOM.EditText)_Form.Items.Item("txtEDate").Specific).Value = calculatedEndDate;

                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        if (!pVal.BeforeAction) // Ensure it runs after selection
                        {
                            if (pVal.ItemUID == "cmbFisY")
                            {
                                SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)_Form.Items.Item("cmbFisY").Specific;
                                string selectedYear = oCombo.Selected.Value;
                                PopulatePostingPeriods(_Form, selectedYear);
                            }
                            else if (pVal.ItemUID == "cmbPosP")
                            {
                                SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)_Form.Items.Item("cmbPosP").Specific;
                                string periodAbsEntry = oCombo.Selected.Value;
                                SetDatesFromPeriod(_Form, periodAbsEntry);
                            }
                        }
                        break;
                }
            }
        }

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
                                //_Form.Items.Item("12").Enabled = true;
                            }
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

        #region Form Default
        internal void FormDefault()
        {
            try
            {
                _Form.Freeze(true);

                _Form.DataBrowser.BrowseBy = "txtDocNo";
                _Form.Settings.Enabled = true;

                _MasterDBDataSource = _Form.DataSources.DBDataSources.Item("@MINTERESTCH");

                PopulateFiscalYears(_Form);

                if (_Form.Mode == BoFormMode.fm_ADD_MODE)
                {
                    SetNextDocNum();
                }
            }
            finally
            {
                _Form.Freeze(false);
            }

        }
        #endregion

        #region Document Number Increment 
        private void SetNextDocNum()
        {
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                // Query to get the max number. Cast to INT to ensure numeric sorting.
                string sql = "SELECT MAX(CAST(IFNULL(\"U_DocNo\", 0) AS INT)) FROM \"@MINTERESTCH\"";
                rs.DoQuery(sql);

                int nextNum = 1;
                if (!rs.EoF)
                {
                    nextNum = Convert.ToInt32(rs.Fields.Item(0).Value) + 1;
                }

                _MasterDBDataSource.SetValue("U_DocNo", 0, nextNum.ToString());
                ((SAPbouiCOM.EditText)_Form.Items.Item("txtDocNo").Specific).Value = nextNum.ToString();
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.StatusBar.SetText("Error generating DocNum: " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }
        #endregion


        private bool IsFirstDayOfNepaliMonth(string engDateStr)
        {
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string sql = $@"SELECT ""U_NepaliDate"" FROM ""@NEPALICALENDER"" WHERE ""U_EnglishDate"" = '{engDateStr}'";
                rs.DoQuery(sql);

                if (!rs.EoF)
                {
                    string nepaliDate = rs.Fields.Item(0).Value.ToString(); // Expected: 2081/05/01
                    return nepaliDate.EndsWith("01");
                }
                return false;
            }
            catch { return false; }
            finally { System.Runtime.InteropServices.Marshal.ReleaseComObject(rs); }
        }

        private string GetNepaliMonthEndDate(string startDateStr)
        {
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                // Get the Nepali Year and Month for the given English Start Date
                string sqlGetMonth = $@"SELECT SUBSTRING(""U_NepaliDate"", 1, 7) as ""YM"" FROM ""@NEPALICALENDER"" WHERE ""U_EnglishDate"" = '{startDateStr}'";
                rs.DoQuery(sqlGetMonth);

                if (rs.EoF) return "";
                string yearMonth = rs.Fields.Item(0).Value.ToString(); // e.g. "2081/05"

                // Find the maximum English date for that Nepali Year/Month
                string sqlGetEnd = $@"SELECT MAX(""U_EnglishDate"") FROM ""@NEPALICALENDER"" WHERE ""U_NepaliDate"" LIKE '{yearMonth}%'";
                rs.DoQuery(sqlGetEnd);

                if (!rs.EoF)
                {
                    DateTime dt = (DateTime)rs.Fields.Item(0).Value;
                    return dt.ToString("yyyyMMdd");
                }
                return "";
            }
            catch { return ""; }
            finally { System.Runtime.InteropServices.Marshal.ReleaseComObject(rs); }
        }


        private bool ValidateFetchCriteria()
        {
            try
            {
                //_Form.DataSources.UserDataSources.Item("txtBCode").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("txtBCode").Specific).Value;

                string bankCode = _MasterDBDataSource.GetValue("U_BankCode", 0).Trim();
                string startDate = _MasterDBDataSource.GetValue("U_StartDate", 0).Trim();

                if (string.IsNullOrEmpty(bankCode))
                {
                    Utilities.Application.SBO_Application.StatusBar.SetText("Bank Code is mandatory.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    _Form.Items.Item("txtBCode").Click();
                    return false;
                }

                if (string.IsNullOrEmpty(startDate))
                {
                    Utilities.Application.SBO_Application.StatusBar.SetText("Start Date is mandatory.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    _Form.Items.Item("txtSDate").Click();
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        //private void FetchData()
        //{
        //    SAPbobsCOM.Recordset rs = null;
        //    SAPbouiCOM.Matrix mtx = (SAPbouiCOM.Matrix)_Form.Items.Item("mtxFtch").Specific;
        //    SAPbouiCOM.DBDataSource dsLines = _Form.DataSources.DBDataSources.Item("@MINTERESTCC");

        //    try
        //    {
        //        _Form.Freeze(true);

        //        // 1. Get User Selection Dates from Header
        //        string bankCode = _MasterDBDataSource.GetValue("U_BankCode", 0).Trim();
        //        string selSDateStr = _MasterDBDataSource.GetValue("U_StartDate", 0).Trim();
        //        string selEDateStr = _MasterDBDataSource.GetValue("U_EndDate", 0).Trim();
        //        string LoanType = _MasterDBDataSource.GetValue("U_LoanType", 0).Trim();

        //        DateTime selSDate = DateTime.ParseExact(selSDateStr, "yyyyMMdd", null);
        //        DateTime selEDate = DateTime.ParseExact(selEDateStr, "yyyyMMdd", null);

        //        //string sql = $@"
        //        //    SELECT 
        //        //        X.""U_LoanType"",
        //        //        X.""U_LoanCode"",
        //        //        X.""Total Interest"",
        //        //        X.""Base Interest"",
        //        //        X.""Premium Interest"",
        //        //        X.""LoanMasterStart"",
        //        //        X.""LoanMasterEnd"",
        //        //        X.""U_BalanceAmount"",
        //        //        MAX(X.""U_ToDate"") AS ""MaxToDate""
        //        //    FROM (
        //        //        SELECT
        //        //            T0.""U_LoanType"", 
        //        //            T0.""U_LoanCode"",
        //        //            T0.""U_Interest"" AS ""Total Interest"",
        //        //            T0.""U_BaseInterest"" AS ""Base Interest"",
        //        //            T0.""U_PremInterest"" AS ""Premium Interest"",
        //        //            T2.""U_StartDate"" AS ""LoanMasterStart"", 
        //        //            T2.""U_EndDate"" AS ""LoanMasterEnd"",
        //        //            T2.""U_BalanceAmount"",
        //        //            T3.""U_ToDate""
        //        //        FROM ""@LOANMASTERC"" T0
        //        //        INNER JOIN ""@LOANMASTERH"" T1 
        //        //            ON T0.""DocEntry"" = T1.""DocEntry""
        //        //        INNER JOIN ""@LOANMASTERDETAILSH"" T2 
        //        //            ON T0.""U_LoanCode"" = T2.""U_LoanCode""
        //        //        INNER JOIN ""@LOANMASTERDETAILSC1"" T3 
        //        //            ON T2.""DocEntry"" = T3.""DocEntry"" 
        //        //        WHERE T1.""U_BankCode"" = '{bankCode}' 
        //        //          AND T2.""U_EndDate"" >= '{selSDateStr}'
        //        //          AND T3.""U_ToDate"" IS NOT NULL
        //        //    ) X
        //        //    GROUP BY
        //        //        X.""U_LoanType"",
        //        //        X.""U_LoanCode"",
        //        //        X.""Total Interest"",
        //        //        X.""Base Interest"",
        //        //        X.""Premium Interest"",
        //        //        X.""LoanMasterStart"",
        //        //        X.""LoanMasterEnd"",
        //        //        X.""U_BalanceAmount"";";

        //        string sql = $@"
        //            SELECT 
        //                T2.""U_LoanCode"",
        //                T3.""U_IntrDateBook"",
        //                T3.""U_NIntrDateBook"",
        //                T3.""U_Remarks"",
        //                T3.""U_InterestPer"",
        //                T3.""U_FromDate"",
        //                T3.""U_ToDate"",
        //                T3.""U_NFromDate"",
        //                T3.""U_NToDate"",
        //                T3.""U_NoDays"",
        //                T3.""U_PrincipalAmt"",
        //                T3.""U_InterestAmt"",
        //                T3.""U_JENo""
        //            FROM ""@LOANMASTERH"" T0
        //            INNER JOIN ""@LOANMASTERC"" T1 
        //                ON T0.""DocEntry"" = T1.""DocEntry""
        //            INNER JOIN ""@LOANMASTERDETAILSH"" T2 
        //                ON T1.""U_LoanCode"" = T2.""U_LoanCode""
        //            INNER JOIN ""@LOANMASTERDETAILSC2"" T3 
        //                ON T2.""DocEntry"" = T3.""DocEntry""
        //            WHERE T0.""U_BankCode"" = '{bankCode}'
        //              AND T1.""U_LoanType"" = '{LoanType}'
        //              AND T2.""U_EndDate"" >= '{selSDateStr}'";

        //        rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
        //        rs.DoQuery(sql);

        //        dsLines.Clear();
        //        int row = 0;

        //        while (!rs.EoF)
        //        {

        //            var checkFrom = rs.Fields.Item("U_FromDate").Value;
        //            var checkTo = rs.Fields.Item("U_ToDate").Value;
        //            string checkIntr = rs.Fields.Item("U_InterestPer").Value.ToString();

        //            bool hasValidDates = (checkFrom is DateTime dtF && dtF.Year > 1900) || (checkTo is DateTime dtT && dtT.Year > 1900);
        //            bool hasInterest = !string.IsNullOrEmpty(checkIntr) && checkIntr != "0";

        //            // If there is no date and no interest percentage, skip this row
        //            if (!hasValidDates && !hasInterest)
        //            {
        //                rs.MoveNext();
        //                continue;
        //            }

        //            dsLines.InsertRecord(row);

        //            // Mapping Recordset to DBDataSource
        //            dsLines.SetValue("U_LoanCode", row, rs.Fields.Item("U_LoanCode").Value.ToString());

        //            var vDateBooked = rs.Fields.Item("U_IntrDateBook").Value;
        //            if (vDateBooked is DateTime dtBooked && dtBooked.Year > 1900)
        //                dsLines.SetValue("U_IntrDateBook", row, dtBooked.ToString("yyyyMMdd"));
        //            else
        //                dsLines.SetValue("U_IntrDateBook", row, "");

        //            dsLines.SetValue("U_NIntrDateBook", row, rs.Fields.Item("U_NIntrDateBook").Value.ToString());


        //            dsLines.SetValue("U_Remarks", row, rs.Fields.Item("U_Remarks").Value.ToString());
        //            dsLines.SetValue("U_InterestPer", row, rs.Fields.Item("U_InterestPer").Value.ToString());

        //            var vFromDate = rs.Fields.Item("U_FromDate").Value;
        //            if (vFromDate is DateTime dtFrom && dtFrom.Year > 1900)
        //                dsLines.SetValue("U_FromDate", row, dtFrom.ToString("yyyyMMdd"));
        //            else
        //                dsLines.SetValue("U_FromDate", row, "");

        //            var vToDate = rs.Fields.Item("U_ToDate").Value;
        //            if (vToDate is DateTime dtTo && dtTo.Year > 1900)
        //                dsLines.SetValue("U_ToDate", row, dtTo.ToString("yyyyMMdd"));
        //            else
        //                dsLines.SetValue("U_ToDate", row, "");

        //            dsLines.SetValue("U_NFromDate", row, rs.Fields.Item("U_NFromDate").Value.ToString());
        //            dsLines.SetValue("U_NToDate", row, rs.Fields.Item("U_NToDate").Value.ToString());
        //            string noDays = rs.Fields.Item("U_NoDays").Value.ToString();
        //            if (noDays == "0" || string.IsNullOrEmpty(noDays))
        //                dsLines.SetValue("U_NoDays", row, "");
        //            else
        //                dsLines.SetValue("U_NoDays", row, noDays);

        //            // Financial Amounts
        //            dsLines.SetValue("U_PrincipalAmt", row, rs.Fields.Item("U_PrincipalAmt").Value.ToString());
        //            dsLines.SetValue("U_InterestAmt", row, rs.Fields.Item("U_InterestAmt").Value.ToString());

        //            dsLines.SetValue("U_JENo", row, rs.Fields.Item("U_JENo").Value.ToString());

        //            row++;
        //            rs.MoveNext();
        //        }

        //        // Load the data into the Matrix UI
        //        mtx.LoadFromDataSource();

        //        if (row == 0)
        //            Utilities.Application.SBO_Application.StatusBar.SetText("No data found for the given criteria.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
        //        else
        //            Utilities.Application.SBO_Application.StatusBar.SetText($"Successfully fetched {row} rows.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

        //    }
        //    catch (Exception ex)
        //    {
        //        Utilities.Application.SBO_Application.StatusBar.SetText("Fetch Error: " + ex.Message);
        //    }
        //    finally
        //    {
        //        _Form.Freeze(false);
        //        if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
        //    }
        //}

        private void FetchData(bool isRefreshCall = false)
        {
            SAPbobsCOM.Recordset rs = null;
            try
            {
                _Form.Freeze(true);

                string bankCode = _MasterDBDataSource.GetValue("U_BankCode", 0).Trim();
                string selSDateStr = _MasterDBDataSource.GetValue("U_StartDate", 0).Trim();
                string selEDateStr = _MasterDBDataSource.GetValue("U_EndDate", 0).Trim();
                string loanType = _MasterDBDataSource.GetValue("U_LoanType", 0).Trim();

                if (string.IsNullOrEmpty(bankCode) || string.IsNullOrEmpty(selSDateStr) || string.IsNullOrEmpty(loanType))
                {
                    Utilities.Application.SBO_Application.StatusBar.SetText("Please select Bank Code, Start Date and Loan Type first.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return;
                }

                rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                if (!isRefreshCall && _Form.Mode == BoFormMode.fm_ADD_MODE)
                {
                    string checkSql = $@"SELECT ""DocEntry"" FROM ""@MINTERESTCH"" WHERE ""U_BankCode"" = '{bankCode}' AND ""U_StartDate"" = '{selSDateStr}' AND ""U_LoanType"" = '{loanType}' ";
                    rs.DoQuery(checkSql);

                    if (!rs.EoF)
                    {
                        // RECORD EXISTS: Load it instead of fetching new data
                        string docNum = rs.Fields.Item(0).Value.ToString();

                        _Form.Mode = BoFormMode.fm_FIND_MODE;
                        _Form.Items.Item("txtDocE").Enabled = true;
                        ((SAPbouiCOM.EditText)_Form.Items.Item("txtDocE").Specific).Value = docNum;
                        _Form.Items.Item("1").Visible = true;
                        _Form.Items.Item("1").Click(); // Simulates clicking the 'Find' button
                        _Form.Items.Item("1").Visible = false;

                        Utilities.Application.SBO_Application.StatusBar.SetText($"Existing record found (DocNum: {docNum}). Loading data...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                        return;
                    }
                }

                // RECORD DOES NOT EXIST: Proceed with pulling from Master Data
                string sql = $@"
                    SELECT 
                        T2.""U_LoanCode"",
                        T3.""U_IntrDateBook"",
                        T3.""U_NIntrDateBook"",
                        T3.""U_Remarks"",
                        T3.""U_InterestPer"",
                        T3.""U_FromDate"",
                        T3.""U_ToDate"",
                        T3.""U_NFromDate"",
                        T3.""U_NToDate"",
                        T3.""U_NoDays"",
                        T3.""U_PrincipalAmtC"",
                        T3.""U_InterestAmtC"",
                        T3.""U_JENo""
                    FROM ""@LOANMASTERH"" T0
                    INNER JOIN ""@LOANMASTERC"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                    INNER JOIN ""@LOANMASTERDETAILSH"" T2 ON T1.""U_LoanCode"" = T2.""U_LoanCode""
                    INNER JOIN ""@LOANMASTERDETAILSC2"" T3 ON T2.""DocEntry"" = T3.""DocEntry""
                    WHERE T0.""U_BankCode"" = '{bankCode}'
                      AND T1.""U_LoanType"" = '{loanType}'
                      AND T2.""U_EndDate"" >= '{selSDateStr}' AND T3.""U_DIntrDateBook"" = '{selEDateStr}'
                   ORDER BY T2.""U_LoanCode"",  (CASE WHEN IFNULL(T3.""U_FromDate"", '') = '' THEN 2 ELSE 1 END),  T3.""U_FromDate"" ASC ";

                rs.DoQuery(sql);

                SAPbouiCOM.Matrix mtx = (SAPbouiCOM.Matrix)_Form.Items.Item("mtxFtch").Specific;
                SAPbouiCOM.DBDataSource dsLines = _Form.DataSources.DBDataSources.Item("@MINTERESTCC");
                dsLines.Clear();
                int row = 0;

                while (!rs.EoF)
                {
                    // Validation: Skip rows where critical dates/interest are missing
                    var checkFrom = rs.Fields.Item("U_FromDate").Value;
                    var checkTo = rs.Fields.Item("U_ToDate").Value;
                    string checkIntr = rs.Fields.Item("U_InterestPer").Value.ToString();

                    bool isSummary = checkIntr.Contains("Total");

                    bool hasValidDates = (checkFrom is DateTime dtF && dtF.Year > 1900) || (checkTo is DateTime dtT && dtT.Year > 1900);
                    bool hasInterest = !string.IsNullOrEmpty(checkIntr) && checkIntr != "0";

                    if (!hasValidDates && !hasInterest)
                    {
                        rs.MoveNext();
                        continue;
                    }

                    dsLines.InsertRecord(row);
                    dsLines.SetValue("U_LoanCode", row, rs.Fields.Item("U_LoanCode").Value.ToString());

                    // Handle Dates (Booked Date)
                    var vDateBooked = rs.Fields.Item("U_IntrDateBook").Value;
                    if (vDateBooked is DateTime dtBooked && dtBooked.Year > 1900)
                        dsLines.SetValue("U_IntrDateBook", row, dtBooked.ToString("yyyyMMdd"));
                    else
                        dsLines.SetValue("U_IntrDateBook", row, "");

                    dsLines.SetValue("U_NIntrDateBook", row, rs.Fields.Item("U_NIntrDateBook").Value.ToString());
                    dsLines.SetValue("U_Remarks", row, rs.Fields.Item("U_Remarks").Value.ToString());
                    dsLines.SetValue("U_InterestPer", row, rs.Fields.Item("U_InterestPer").Value.ToString());

                    // Handle From/To Dates
                    if (checkFrom is DateTime dtFrom && dtFrom.Year > 1900)
                        dsLines.SetValue("U_FromDate", row, dtFrom.ToString("yyyyMMdd"));
                    else
                        dsLines.SetValue("U_FromDate", row, "");

                    if (checkTo is DateTime dtTo && dtTo.Year > 1900)
                        dsLines.SetValue("U_ToDate", row, dtTo.ToString("yyyyMMdd"));
                    else
                        dsLines.SetValue("U_ToDate", row, "");

                    dsLines.SetValue("U_NFromDate", row, rs.Fields.Item("U_NFromDate").Value.ToString());
                    dsLines.SetValue("U_NToDate", row, rs.Fields.Item("U_NToDate").Value.ToString());

                    string noDays = rs.Fields.Item("U_NoDays").Value.ToString();
                    dsLines.SetValue("U_NoDays", row, (noDays == "0" || string.IsNullOrEmpty(noDays)) ? "" : noDays);

                    dsLines.SetValue("U_PrincipalAmtC", row, rs.Fields.Item("U_PrincipalAmtC").Value.ToString());
                    dsLines.SetValue("U_InterestAmtC", row, rs.Fields.Item("U_InterestAmtC").Value.ToString());
                    dsLines.SetValue("U_JENo", row, rs.Fields.Item("U_JENo").Value.ToString());

                    row++;
                    rs.MoveNext();
                }

                mtx.LoadFromDataSource();
                if (isRefreshCall)
                    Utilities.Application.SBO_Application.StatusBar.SetText("Matrix updated with calculated values.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.StatusBar.SetText("Fetch Error: " + ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }

        private bool ValidateSlabGapsLineWise()
        {
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            SAPbouiCOM.DBDataSource dsLines = _Form.DataSources.DBDataSources.Item("@MINTERESTCC");
            Dictionary<string, DateTime> internalTracker = new Dictionary<string, DateTime>();

            try
            {
                for (int i = 0; i < dsLines.Size; i++)
                {
                    string isChecked = dsLines.GetValue("U_Check", i).Trim();
                    if (isChecked != "Y") continue;


                    string loanCode = dsLines.GetValue("U_LoanNo", i).Trim();
                    string rowSDateStr = dsLines.GetValue("U_StartDate", i).Trim();
                    string rowEDateStr = dsLines.GetValue("U_EndDate", i).Trim();
                    if (string.IsNullOrEmpty(loanCode)) continue;

                    DateTime rowSDate = DateTime.ParseExact(rowSDateStr, "yyyyMMdd", null);
                    DateTime rowEDate = DateTime.ParseExact(rowEDateStr, "yyyyMMdd", null);

                    // 1. Fetch Master Details (Start, End, and Last Slab)
                    string sql = $@"
                SELECT 
                    T0.""U_StartDate"" as ""MasterStart"",
                    T0.""U_EndDate"" as ""MasterEnd"",
                    (SELECT MAX(T1.""U_ToDate"") FROM ""@LOANMASTERDETAILSC1"" T1 WHERE T1.""DocEntry"" = T0.""DocEntry"") as ""LastSlabTo""
                FROM ""@LOANMASTERDETAILSH"" T0
                WHERE T0.""U_LoanCode"" = '{loanCode}'";

                    rs.DoQuery(sql);
                    if (rs.EoF) continue;

                    DateTime masterStart = (DateTime)rs.Fields.Item("MasterStart").Value;
                    DateTime masterEnd = (DateTime)rs.Fields.Item("MasterEnd").Value;
                    object lastSlabVal = rs.Fields.Item("LastSlabTo").Value;

                    // ==========================================================
                    // VALIDATION A: MATURITY CHECK
                    // ==========================================================
                    // Check if the calculation period exceeds the Loan Maturity Date
                    if (rowSDate > masterEnd)
                    {
                        Utilities.Application.SBO_Application.MessageBox($"Loan {loanCode} has already matured on {masterEnd:yyyy-MM-dd}. Cannot calculate interest after this date.");
                        return false;
                    }

                    if (rowEDate > masterEnd)
                    {
                        Utilities.Application.SBO_Application.MessageBox($"The End Date for Loan {loanCode} ({rowEDate:yyyy-MM-dd}) exceeds the Loan Maturity Date ({masterEnd:yyyy-MM-dd}).");
                        return false;
                    }

                    // ==========================================================
                    // VALIDATION B: GAP DETECTION
                    // ==========================================================
                    DateTime expectedSDate;
                    if (internalTracker.ContainsKey(loanCode))
                    {
                        // Check gap within the current matrix
                        expectedSDate = internalTracker[loanCode].AddDays(1);
                        if (rowSDate != expectedSDate)
                        {
                            Utilities.Application.SBO_Application.MessageBox($"Matrix Gap for {loanCode}. Row {i + 1} starts {rowSDate:yyyy-MM-dd}, but previous row ends {internalTracker[loanCode]:yyyy-MM-dd}.");
                            return false;
                        }
                    }
                    else
                    {
                        // Check gap against database history
                        expectedSDate = (lastSlabVal != null && lastSlabVal != DBNull.Value)
                                         ? ((DateTime)lastSlabVal).AddDays(1)
                                         : masterStart;

                        if (rowSDate != expectedSDate)
                        {
                            Utilities.Application.SBO_Application.MessageBox($"Rate History Gap for {loanCode}.\nExpected: {expectedSDate:yyyy-MM-dd}\nFound: {rowSDate:yyyy-MM-dd}");
                            return false;
                        }
                    }

                    internalTracker[loanCode] = rowEDate;
                }
                return true;
            }
            finally { System.Runtime.InteropServices.Marshal.ReleaseComObject(rs); }
        }

        private bool PostSlabsToLoanMasters()
        {
            SAPbobsCOM.GeneralService oGenService = Utilities.Application.Company.GetCompanyService().GetGeneralService("LOANMASTERDETAILS");
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            SAPbouiCOM.DBDataSource dsLines = _Form.DataSources.DBDataSources.Item("@MINTERESTCC");

            try
            {
                _Form.Freeze(true);
                // No need to FlushToDataSource here as it's done in Item_Event

                for (int i = 0; i < dsLines.Size; i++)
                {
                    string isChecked = dsLines.GetValue("U_Check", i).Trim();
                    if (isChecked != "Y") continue;

                    string loanCode = dsLines.GetValue("U_LoanNo", i).Trim();
                    if (string.IsNullOrEmpty(loanCode)) continue;

                    try
                    {
                        DateTime rSDate = DateTime.ParseExact(dsLines.GetValue("U_StartDate", i), "yyyyMMdd", null);
                        DateTime rEDate = DateTime.ParseExact(dsLines.GetValue("U_EndDate", i), "yyyyMMdd", null);

                        // 1. Find DocEntry of the target Loan Master
                        rs.DoQuery($"SELECT \"DocEntry\" FROM \"@LOANMASTERDETAILSH\" WHERE \"U_LoanCode\" = '{loanCode}'");
                        if (rs.EoF) throw new Exception($"Loan Master not found for Code: {loanCode}");

                        int mDocEntry = Convert.ToInt32(rs.Fields.Item(0).Value);

                        // 2. Get the Master UDO Record
                        SAPbobsCOM.GeneralDataParams oParams = (SAPbobsCOM.GeneralDataParams)oGenService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                        oParams.SetProperty("DocEntry", mDocEntry);
                        SAPbobsCOM.GeneralData oMaster = oGenService.GetByParams(oParams);

                        // 3. Child Table C1
                        SAPbobsCOM.GeneralDataCollection oSlabs = oMaster.Child("LOANMASTERDETAILSC1");

                        // 4. Duplicate Check
                        bool slabAlreadyExists = false;
                        for (int j = 0; j < oSlabs.Count; j++)
                        {
                            object existingFrom = oSlabs.Item(j).GetProperty("U_FromDate");
                            if (existingFrom != null && (DateTime)existingFrom == rSDate)
                            {
                                slabAlreadyExists = true;
                                break;
                            }
                        }

                        if (!slabAlreadyExists)
                        {
                            SAPbobsCOM.GeneralData oRowToUpdate = null;

                            // Handle Ghost Row logic (reusing empty rows in UI)
                            if (oSlabs.Count > 0)
                            {
                                SAPbobsCOM.GeneralData oLastRow = oSlabs.Item(oSlabs.Count - 1);
                                object lastFrom = oLastRow.GetProperty("U_FromDate");
                                if (lastFrom == null || (DateTime)lastFrom == DateTime.MinValue || ((DateTime)lastFrom).Year < 1920)
                                {
                                    oRowToUpdate = oLastRow;
                                }
                            }

                            if (oRowToUpdate == null) oRowToUpdate = oSlabs.Add();

                            oRowToUpdate.SetProperty("U_FromDate", rSDate);
                            oRowToUpdate.SetProperty("U_ToDate", rEDate);
                            oRowToUpdate.SetProperty("U_BaseInterestPer", dsLines.GetValue("U_BaseRate", i));
                            oRowToUpdate.SetProperty("U_PremInterestPer", dsLines.GetValue("U_PremRate", i));
                            oRowToUpdate.SetProperty("U_InterestPer", dsLines.GetValue("U_TotalRate", i));
                            oRowToUpdate.SetProperty("U_NoDays", dsLines.GetValue("U_NoDays", i));
                            oRowToUpdate.SetProperty("U_NFromDate", GetNepaliDate(rSDate));
                            oRowToUpdate.SetProperty("U_NToDate", GetNepaliDate(rEDate));

                            oSlabs.Add(); // Keep Master UI healthy
                            oGenService.Update(oMaster);
                        }
                    }
                    catch (Exception rowEx)
                    {
                        // If a row fails, we stop the whole ADD process
                        Utilities.Application.SBO_Application.MessageBox($"Sync failed for Loan {loanCode}: {rowEx.Message}");
                        return false;
                    }
                }
                return true; // Everything succeeded
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.MessageBox("Critical Sync Error: " + ex.Message);
                return false;
            }
            finally
            {
                _Form.Freeze(false);
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }

        // Helper method needed for the synchronization
        private string GetNepaliDate(DateTime engDate)
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

        private bool CheckAndLoadExistingRecord()
        {
            if (_Form.Mode != BoFormMode.fm_ADD_MODE) return false;

            string bankCode = _MasterDBDataSource.GetValue("U_BankCode", 0).Trim();
            string startDate = _MasterDBDataSource.GetValue("U_StartDate", 0).Trim();

            // Only proceed if both fields are filled
            if (string.IsNullOrEmpty(bankCode) || string.IsNullOrEmpty(startDate)) return false;

            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                string sql = $@"SELECT ""DocEntry"" FROM ""@MINTERESTCH"" WHERE ""U_BankCode"" = '{bankCode}' AND ""U_StartDate"" = '{startDate}'";
                rs.DoQuery(sql);

                if (!rs.EoF)
                {
                    string docEntry = rs.Fields.Item(0).Value.ToString();

                    // Switch to Find Mode

                    System.Windows.Forms.Application.DoEvents();
                    System.Threading.Thread.Sleep(100);

                    _Form.Freeze(true);
                    try
                    {
                        // 1. Switch to Find Mode
                        _Form.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

                        // 2. Enable the DocEntry field (txtDocE) so we can write to it
                        SAPbouiCOM.EditText oEditDocE = (SAPbouiCOM.EditText)_Form.Items.Item("txtDocE").Specific;
                        oEditDocE.Item.Enabled = true;
                        oEditDocE.Value = docEntry;

                        // 3. Click the Find button
                        _Form.Items.Item("1").Click();
                    }
                    finally
                    {
                        _Form.Freeze(false);
                    }

                    Utilities.Application.SBO_Application.StatusBar.SetText($"Existing record found for this Bank and Date.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                    return true;
                }
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.StatusBar.SetText("Error checking existing records: " + ex.Message);
            }
            return false;
        }


        private void CalculateInterestForAllRows1()
        {
            SAPbobsCOM.GeneralService oGenService = Utilities.Application.Company.GetCompanyService().GetGeneralService("LOANMASTERDETAILS");
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                _Form.Freeze(true);
                Utilities.Application.SBO_Application.StatusBar.SetText("Starting Interest Calculation...", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);

                string bankCode = _MasterDBDataSource.GetValue("U_BankCode", 0).Trim();
                string loanType = _MasterDBDataSource.GetValue("U_LoanType", 0).Trim();
                string monDocNum = _MasterDBDataSource.GetValue("U_DocNo", 0).Trim();
                string monSDateStr = _MasterDBDataSource.GetValue("U_StartDate", 0).Trim();
                string monEDateStr = _MasterDBDataSource.GetValue("U_EndDate", 0).Trim();

                DateTime monEDate = DateTime.ParseExact(monEDateStr, "yyyyMMdd", null);
                DateTime monSDate = DateTime.ParseExact(monSDateStr, "yyyyMMdd", null);

                // Fetch eligible loans
                string sql = $@"SELECT DISTINCT T2.""U_LoanCode"" FROM ""@LOANMASTERH"" T0 
                        INNER JOIN ""@LOANMASTERC"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                        INNER JOIN ""@LOANMASTERDETAILSH"" T2 ON T1.""U_LoanCode"" = T2.""U_LoanCode""
                        WHERE T0.""U_BankCode"" = '{bankCode}' AND T1.""U_LoanType"" = '{loanType}' AND T2.""U_EndDate"" >= '{monSDateStr}'";
                rs.DoQuery(sql);

                while (!rs.EoF)
                {
                    string loanCode = rs.Fields.Item("U_LoanCode").Value.ToString();
                    SAPbobsCOM.Recordset rsEntry = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    rsEntry.DoQuery($@"SELECT ""DocEntry"" FROM ""@LOANMASTERDETAILSH"" WHERE ""U_LoanCode"" = '{loanCode}'");

                    if (!rsEntry.EoF)
                    {
                        int mDocEntry = Convert.ToInt32(rsEntry.Fields.Item(0).Value);
                        SAPbobsCOM.GeneralDataParams oParams = (SAPbobsCOM.GeneralDataParams)oGenService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                        oParams.SetProperty("DocEntry", mDocEntry);

                        SAPbobsCOM.GeneralData oMaster = oGenService.GetByParams(oParams);
                        SAPbobsCOM.GeneralDataCollection oProvisions = oMaster.Child("LOANMASTERDETAILSC2");

                        double alreadyBookedInMonth = 0;

                        // CLEANUP: Remove details AND summaries matching this Book Date (if no JE exists)
                        for (int row = oProvisions.Count - 1; row >= 0; row--)
                        {
                            var rowData = oProvisions.Item(row);
                            object rBookDateVal = rowData.GetProperty("U_DIntrDateBook");
                            string jeNo = rowData.GetProperty("U_JENo")?.ToString();
                            string interestPer = rowData.GetProperty("U_InterestPer")?.ToString() ?? "";

                            // Only look at rows for the selected month
                            if (rBookDateVal is DateTime dtBook && dtBook.Date == monEDate.Date)
                            {
                                if (!string.IsNullOrEmpty(jeNo))
                                {
                                    // IT HAS A JE: DO NOT TOUCH IT. Just add to our counter.
                                    // We only count Detail rows, not summary rows (though JE rows shouldn't be summaries)
                                    if (!interestPer.Contains("Total"))
                                    {
                                        alreadyBookedInMonth += Convert.ToDouble(rowData.GetProperty("U_InterestAmt"));
                                    }
                                }
                                else
                                {
                                    // NO JE: This is either a Summary Row or an unbooked calculation. 
                                    // DELETE IT so we can recreate it accurately.
                                    oProvisions.Remove(row);
                                }
                            }
                        }

                        // Calculation logic with Summary Row insertion
                        ProcessMonthlySegments(oProvisions, loanCode, monSDate, monEDate, monDocNum, alreadyBookedInMonth);
                        oGenService.Update(oMaster);
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rsEntry);
                    rs.MoveNext();
                }
                FetchData(true); // Refresh Matrix
            }
            catch (Exception ex) { Utilities.Application.SBO_Application.MessageBox("Calc Error: " + ex.Message); }
            finally { _Form.Freeze(false); if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs); }
        }


        private void ProcessMonthlySegments1(SAPbobsCOM.GeneralDataCollection oProvisions, string loanCode, DateTime monStart, DateTime monEnd, string monDocNum, double startingTotal)
        {
            SAPbobsCOM.Recordset rsSlabs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            //double loanMonthlyTotal = 0; // The subtotal for the summary row
            double loanMonthlyTotal = startingTotal;


            try
            {
                string sqlSlabs = $@"SELECT T1.""U_FromDate"", T1.""U_ToDate"", T1.""U_InterestPer"", T0.""U_BalanceAmountC""
                            FROM ""@LOANMASTERDETAILSH"" T0 INNER JOIN ""@LOANMASTERDETAILSC1"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                            WHERE T0.""U_LoanCode"" = '{loanCode}' 
                            AND T1.""U_FromDate"" <= '{monEnd:yyyyMMdd}' AND T1.""U_ToDate"" >= '{monStart:yyyyMMdd}' AND T1.""U_Status"" = 'O' ";
                rsSlabs.DoQuery(sqlSlabs);

                while (!rsSlabs.EoF)
                {
                    // ... (Your existing logic to get slab variables) ...
                    double rate = Convert.ToDouble(rsSlabs.Fields.Item("U_InterestPer").Value);
                    double defBal = Convert.ToDouble(rsSlabs.Fields.Item("U_BalanceAmountC").Value);
                    DateTime slabF = (DateTime)rsSlabs.Fields.Item("U_FromDate").Value;
                    DateTime slabT = (DateTime)rsSlabs.Fields.Item("U_ToDate").Value;
                    DateTime currentStart = slabF < monStart ? monStart : slabF;
                    DateTime segmentLimit = slabT > monEnd ? monEnd : slabT;

                    while (currentStart <= segmentLimit)
                    {
                        DateTime currentEnd = GetNextSegmentEnd(loanCode, currentStart, segmentLimit);
                        double principal = GetPrincipalForDate(loanCode, currentStart, defBal);
                        int days = (currentEnd - currentStart).Days + 1;

                        double interest = Math.Round((principal * rate / 100.0) * (days / 365.0), 2);

                        // Add Detail Row
                        var newRow = oProvisions.Add();
                        newRow.SetProperty("U_FromDate", currentStart);
                        newRow.SetProperty("U_ToDate", currentEnd);
                        newRow.SetProperty("U_InterestPer", rate.ToString());
                        newRow.SetProperty("U_PrincipalAmtC", principal.ToString());
                        newRow.SetProperty("U_InterestAmtC", interest.ToString());
                        newRow.SetProperty("U_NoDays", days.ToString());
                        newRow.SetProperty("U_IntrDateBook", monEnd);
                        newRow.SetProperty("U_DIntrDateBook", monEnd);
                        newRow.SetProperty("U_NIntrDateBook", GetNepaliDate(monEnd));
                        newRow.SetProperty("U_Remarks", "Due to Interest Period Change");
                        newRow.SetProperty("U_MonDocNum", monDocNum);

                        loanMonthlyTotal += interest;
                        currentStart = currentEnd.AddDays(1);
                    }
                    rsSlabs.MoveNext();
                }

                // PHYSICAL SUMMARY ROW SAVE
                if (loanMonthlyTotal > 0)
                {
                    string nepDate = GetNepaliDate(monEnd);
                    string nepYM = nepDate.Length >= 7 ? nepDate.Substring(0, 7) : nepDate;

                    var summaryRow = oProvisions.Add();
                    summaryRow.SetProperty("U_InterestPer", $"Monthly Total Interest for (BS) {nepYM}");
                    //summaryRow.SetProperty("U_InterestAmt", Math.Round(loanMonthlyTotal, 2).ToString());
                    //summaryRow.SetProperty("U_IntrDateBook", monEnd); // Essential for cleanup
                    summaryRow.SetProperty("U_DIntrDateBook", monEnd);
                    summaryRow.SetProperty("U_MonDocNum", monDocNum);
                    //summaryRow.SetProperty("U_NIntrDateBook", nepDate);
                    summaryRow.SetProperty("U_Remarks", Math.Round(loanMonthlyTotal, 2).ToString());
                }
            }
            finally { System.Runtime.InteropServices.Marshal.ReleaseComObject(rsSlabs); }
        }

        private DateTime GetNextSegmentEnd(string loanCode, DateTime start, DateTime limit)
        {
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                string sql = $@"SELECT TOP 1 T1.""U_PaymentDate"" FROM ""@LOANMASTERDETAILSC"" T1 
                        INNER JOIN ""@LOANMASTERDETAILSH"" T0 ON T0.""DocEntry"" = T1.""DocEntry""
                        WHERE T0.""U_LoanCode"" = '{loanCode}' AND T1.""U_PaymentDate"" > '{start:yyyyMMdd}' AND T1.""U_PaymentDate"" <= '{limit:yyyyMMdd}'
                        ORDER BY T1.""U_PaymentDate"" ASC";
                rs.DoQuery(sql);
                // return (!rs.EoF) ? ((DateTime)rs.Fields.Item(0).Value).AddDays(-1) : limit;
                return (!rs.EoF) ? ((DateTime)rs.Fields.Item(0).Value) : limit;
            }
            finally { System.Runtime.InteropServices.Marshal.ReleaseComObject(rs); }
        }

        private double GetPrincipalForDate(string loanCode, DateTime date, double defBal)
        {
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                //string sql = $@"SELECT TOP 1 T1.""U_BalanceAmountC"" 
                //        FROM ""@LOANMASTERDETAILSH"" T0 INNER JOIN ""@LOANMASTERDETAILSC"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                //        WHERE T0.""U_LoanCode"" = '{loanCode}' AND T1.""U_PaymentDate"" <= '{date:yyyyMMdd}'
                //        ORDER BY T1.""U_PaymentDate"" DESC, T1.""LineId"" DESC";
                //rs.DoQuery(sql);
                //if (!rs.EoF) return Convert.ToDouble(rs.Fields.Item(0).Value);

                //string sqlFuture = $@"SELECT TOP 1 T1.""U_UnPaidAmountC"" 
                //             FROM ""@LOANMASTERDETAILSH"" T0 
                //             INNER JOIN ""@LOANMASTERDETAILSC"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                //             WHERE T0.""U_LoanCode"" = '{loanCode}' 
                //               AND T1.""U_PaymentDate"" > '{date:yyyyMMdd}' 
                //             ORDER BY T1.""U_PaymentDate"" ASC";


                string sql = $@"SELECT TOP 1 T1.""U_BalanceAmountC"" 
                        FROM ""@LOANMASTERDETAILSH"" T0 INNER JOIN ""@LOANMASTERDETAILSC"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                        WHERE T0.""U_LoanCode"" = '{loanCode}' AND T1.""U_PaymentDate"" < '{date:yyyyMMdd}'
                        ORDER BY T1.""U_PaymentDate"" DESC, T1.""LineId"" DESC";
                rs.DoQuery(sql);
                if (!rs.EoF) return Convert.ToDouble(rs.Fields.Item(0).Value);

                string sqlFuture = $@"SELECT TOP 1 T1.""U_UnPaidAmountC"" 
                             FROM ""@LOANMASTERDETAILSH"" T0 
                             INNER JOIN ""@LOANMASTERDETAILSC"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                             WHERE T0.""U_LoanCode"" = '{loanCode}' 
                               AND T1.""U_PaymentDate"" >= '{date:yyyyMMdd}' 
                             ORDER BY T1.""U_PaymentDate"" ASC";

                rs.DoQuery(sqlFuture);
                return (!rs.EoF) ? Convert.ToDouble(rs.Fields.Item(0).Value) : defBal;
            }
            finally { System.Runtime.InteropServices.Marshal.ReleaseComObject(rs); }
        }


        private void BookInterestBatch()
        {
            SAPbobsCOM.Recordset rs = null;
            SAPbobsCOM.JournalEntries oJE = null;
            SAPbouiCOM.Matrix mtx = (SAPbouiCOM.Matrix)_Form.Items.Item("mtxFtch").Specific;
            SAPbouiCOM.DBDataSource dsLines = _Form.DataSources.DBDataSources.Item("@MINTERESTCC");

            try
            {
                if (!ValidatePreviousMonthBooked())
                {
                    return;
                }

                // 1. ALWAYS RECALCULATE FIRST to handle recent repayments
                CalculateInterestForAllRows();
                mtx.FlushToDataSource(); // Sync matrix selection state

                _Form.Freeze(true);
                Utilities.Application.SBO_Application.StatusBar.SetText("Starting Batch Booking...", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);

                string currentNepDate = GetNepaliDate(DateTime.Today);
                string currentNepMonthKey = currentNepDate.Substring(0, 7);

                int successCount = 0;

                for (int i = 0; i < dsLines.Size; i++)
                {
                    // Only process checked rows
                    //if (dsLines.GetValue("U_Check", i).Trim() != "Y") continue;

                    string loanCode = dsLines.GetValue("U_LoanCode", i).Trim();
                    string interestAmtStr = dsLines.GetValue("U_InterestAmtC", i).Trim();
                    string existingJE = dsLines.GetValue("U_JENo", i).Trim();
                    string fromDateStr = dsLines.GetValue("U_FromDate", i).Trim();
                    string toDateStr = dsLines.GetValue("U_ToDate", i).Trim();
                    string bookDateStr = dsLines.GetValue("U_IntrDateBook", i).Trim();

                    // SKIP: Summary rows (no From Date), zero amounts, or already booked
                    if (string.IsNullOrEmpty(fromDateStr) || string.IsNullOrEmpty(interestAmtStr) || !string.IsNullOrEmpty(existingJE))
                        continue;

                    double interestAmt = Convert.ToDouble(interestAmtStr);
                    if (interestAmt <= 0) continue;

                    // PROTECTION: Only book previous months
                    DateTime dtBookDate = DateTime.ParseExact(bookDateStr, "yyyyMMdd", null);
                    string rowNepMonthKey = GetNepaliDate(dtBookDate).Substring(0, 7);
                    if (string.Compare(rowNepMonthKey, currentNepMonthKey) >= 0) continue;

                    // 2. FETCH GL ACCOUNTS for this specific Loan
                    rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    string sqlGL = $@"
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

                    rs.DoQuery(sqlGL);
                    if (rs.EoF) continue;

                    string debitAcc = rs.Fields.Item("Interest").Value.ToString();
                    string creditAcc = rs.Fields.Item("Provision").Value.ToString();
                    int dimension = Convert.ToInt32(rs.Fields.Item("DimCode").Value);

                    // 3. CREATE JE
                    oJE = (SAPbobsCOM.JournalEntries)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.oJournalEntries);
                    oJE.ReferenceDate = dtBookDate;
                    oJE.Memo = $"Interest Provision - {loanCode} - {rowNepMonthKey}";

                    // Debit Interest / Credit Provision
                    AddJELine(oJE, debitAcc, interestAmt, 0, 1, loanCode, dimension, false);
                    AddJELine(oJE, creditAcc, 0, interestAmt, 1, loanCode, dimension, true);

                    if (oJE.Add() == 0)
                    {
                        string newJeKey = Utilities.Application.Company.GetNewObjectKey();

                        // 4. UPDATE MASTER DATA TABLE (@LOANMASTERDETAILSC2)
                        UpdateMasterWithJE(loanCode, fromDateStr, toDateStr, bookDateStr, newJeKey);

                        // 5. UPDATE LOCAL DATASOURCE
                        dsLines.SetValue("U_JENo", i, newJeKey);
                        successCount++;
                    }
                    else
                    {
                        int errCode;
                        string errMsg;
                        Utilities.Application.Company.GetLastError(out errCode, out errMsg);

                        // Display the error in the SAP Status Bar (Red Bar)
                        Utilities.Application.SBO_Application.StatusBar.SetText(
                            $"Journal Entry Error ({errCode}): {errMsg}",
                            SAPbouiCOM.BoMessageTime.bmt_Medium,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Error
                        );
                    }

                    if (oJE != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oJE);
                    if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                }

                mtx.LoadFromDataSource();
                _Form.Items.Item("1").Visible = true;
                _Form.Mode = BoFormMode.fm_UPDATE_MODE;
                _Form.Items.Item("1").Click(); // Save JE numbers to the Monthly Calculation doc
                _Form.Items.Item("1").Visible = false;

                Utilities.Application.SBO_Application.StatusBar.SetText($"Successfully booked {successCount} JEs.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex) { Utilities.Application.SBO_Application.MessageBox(ex.Message); }
            finally { _Form.Freeze(false); }
        }

        // Helper to update the actual Loan Master record
        private void UpdateMasterWithJE(string loanCode, string fromDate, string toDate, string bookDate, string jeNo)
        {
            SAPbobsCOM.GeneralService oGenService = Utilities.Application.Company.GetCompanyService().GetGeneralService("LOANMASTERDETAILS");
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);

            rs.DoQuery($@"SELECT ""DocEntry"" FROM ""@LOANMASTERDETAILSH"" WHERE ""U_LoanCode"" = '{loanCode}'");
            if (rs.EoF) return;

            int docEntry = Convert.ToInt32(rs.Fields.Item(0).Value);
            SAPbobsCOM.GeneralDataParams oParams = (SAPbobsCOM.GeneralDataParams)oGenService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
            oParams.SetProperty("DocEntry", docEntry);



            SAPbobsCOM.GeneralData oMaster = oGenService.GetByParams(oParams);


            //SAPbobsCOM.GeneralDataCollection oProvisions = oMaster.Child("LOANMASTERDETAILSC2");

            //for (int i = 0; i < oProvisions.Count; i++)
            //{
            //    var row = oProvisions.Item(i);
            //    DateTime dtFrom = (DateTime)row.GetProperty("U_FromDate");
            //    DateTime dtBook = (DateTime)row.GetProperty("U_IntrDateBook");

            //    if (dtFrom.ToString("yyyyMMdd") == fromDate && dtBook.ToString("yyyyMMdd") == bookDate)
            //    {
            //        row.SetProperty("U_JENo", jeNo);
            //        break;
            //    }
            //}

            DateTime dtFrom = DateTime.ParseExact(fromDate, "yyyyMMdd", null);
            DateTime dtTo = DateTime.ParseExact(toDate, "yyyyMMdd", null);
            DateTime dtBook = DateTime.ParseExact(bookDate, "yyyyMMdd", null);

            //Update Status in @LOANMASTERDETAILSC1 
            SAPbobsCOM.GeneralDataCollection oSlabs = oMaster.Child("LOANMASTERDETAILSC1");
            for (int i = 0; i < oSlabs.Count; i++)
            {
                var slabRow = oSlabs.Item(i);
                object sFromVal = slabRow.GetProperty("U_FromDate");
                object sToVal = slabRow.GetProperty("U_ToDate");

                if (sFromVal is DateTime sFrom && sToVal is DateTime sTo)
                {
                    // Check if the current booking segment falls within this slab's range
                    if (dtFrom >= sFrom.Date && dtTo <= sTo.Date)
                    {
                        slabRow.SetProperty("U_Status", "C"); // Set Status to Closed
                    }
                }
            }

            SAPbobsCOM.GeneralDataCollection oProvisions = oMaster.Child("LOANMASTERDETAILSC2");
            for (int i = 0; i < oProvisions.Count; i++)
            {
                var provRow = oProvisions.Item(i);
                object rFromVal = provRow.GetProperty("U_FromDate");
                object rBookVal = provRow.GetProperty("U_IntrDateBook");

                if (rFromVal is DateTime rFrom && rBookVal is DateTime rBook)
                {
                    if (rFrom.Date == dtFrom.Date && rBook.Date == dtBook.Date)
                    {
                        provRow.SetProperty("U_JENo", jeNo);
                    }
                }
            }

            oGenService.Update(oMaster);
        }


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

        private void CalculateInterestForAllRows()
        {
            SAPbobsCOM.GeneralService oGenService = Utilities.Application.Company.GetCompanyService().GetGeneralService("LOANMASTERDETAILS");
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                _Form.Freeze(true);
                Utilities.Application.SBO_Application.StatusBar.SetText("Starting Interest Calculation...", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);

                string bankCode = _MasterDBDataSource.GetValue("U_BankCode", 0).Trim();
                string loanType = _MasterDBDataSource.GetValue("U_LoanType", 0).Trim();
                string monDocNum = _MasterDBDataSource.GetValue("U_DocNo", 0).Trim();
                string monSDateStr = _MasterDBDataSource.GetValue("U_StartDate", 0).Trim();
                string monEDateStr = _MasterDBDataSource.GetValue("U_EndDate", 0).Trim();

                DateTime monEDate = DateTime.ParseExact(monEDateStr, "yyyyMMdd", null);
                DateTime monSDate = DateTime.ParseExact(monSDateStr, "yyyyMMdd", null);

                // Fetch eligible loans
                string sql = $@"SELECT DISTINCT T2.""U_LoanCode"" FROM ""@LOANMASTERH"" T0 
                        INNER JOIN ""@LOANMASTERC"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                        INNER JOIN ""@LOANMASTERDETAILSH"" T2 ON T1.""U_LoanCode"" = T2.""U_LoanCode""
                        WHERE T0.""U_BankCode"" = '{bankCode}' AND T1.""U_LoanType"" = '{loanType}' AND T2.""U_EndDate"" >= '{monSDateStr}'
                        AND T2.""U_BalanceAmountC"" > 0 ";
                rs.DoQuery(sql);

                while (!rs.EoF)
                {
                    string loanCode = rs.Fields.Item("U_LoanCode").Value.ToString();
                    SAPbobsCOM.Recordset rsEntry = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    rsEntry.DoQuery($@"SELECT ""DocEntry"" FROM ""@LOANMASTERDETAILSH"" WHERE ""U_LoanCode"" = '{loanCode}'");

                    if (!rsEntry.EoF)
                    {
                        int mDocEntry = Convert.ToInt32(rsEntry.Fields.Item(0).Value);
                        SAPbobsCOM.GeneralDataParams oParams = (SAPbobsCOM.GeneralDataParams)oGenService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                        oParams.SetProperty("DocEntry", mDocEntry);

                        SAPbobsCOM.GeneralData oMaster = oGenService.GetByParams(oParams);
                        SAPbobsCOM.GeneralDataCollection oProvisions = oMaster.Child("LOANMASTERDETAILSC2");

                        List<ProvisionRowData> futureRows = new List<ProvisionRowData>();
                        double alreadyBookedInMonth = 0;

                        // 1. BACKUP FUTURE ROWS & REMOVE CURRENT UNBOOKED ROWS
                        // We iterate backwards to maintain index integrity while removing
                        for (int i = oProvisions.Count - 1; i >= 0; i--)
                        {
                            var row = oProvisions.Item(i);
                            DateTime rBookDate = (DateTime)row.GetProperty("U_DIntrDateBook");
                            string jeNo = row.GetProperty("U_JENo")?.ToString();
                            string interestPer = row.GetProperty("U_InterestPer")?.ToString() ?? "";

                            if (rBookDate.Date > monEDate.Date)
                            {
                                // Save future month data
                                futureRows.Add(ExtractRowData(row));
                                oProvisions.Remove(i);
                            }
                            else if (rBookDate.Date == monEDate.Date)
                            {
                                if (!string.IsNullOrEmpty(jeNo))
                                {
                                    // If it has a JE, keep it and add to counter
                                    if (!interestPer.Contains("Total"))
                                        alreadyBookedInMonth += Convert.ToDouble(row.GetProperty("U_InterestAmtC"));
                                }
                                else
                                {
                                    // No JE? Remove it so we can re-insert details + summary in order
                                    oProvisions.Remove(i);
                                }
                            }
                        }

                        // 2. PROCESS CURRENT MONTH (Details then Summary)
                        // This appends lines right after the previous month's rows
                        ProcessMonthlySegments(oProvisions, loanCode, monSDate, monEDate, monDocNum, alreadyBookedInMonth);

                        // 3. RESTORE FUTURE ROWS
                        // Reverse the list because we added them while looping backwards
                        futureRows.Reverse();
                        foreach (var fRow in futureRows)
                        {
                            InsertRowData(oProvisions.Add(), fRow);
                        }

                        oGenService.Update(oMaster);
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rsEntry);
                    rs.MoveNext();
                }
                FetchData(true);

                if (_Form.Mode == BoFormMode.fm_OK_MODE)
                {
                    _Form.Items.Item("1").Visible = true;
                    _Form.Mode = BoFormMode.fm_UPDATE_MODE;
                    _Form.Items.Item("1").Visible = false;
                }


                if (_Form.Mode == BoFormMode.fm_ADD_MODE || _Form.Mode == BoFormMode.fm_UPDATE_MODE)
                {
                    SAPbouiCOM.Matrix mtx = (SAPbouiCOM.Matrix)_Form.Items.Item("mtxFtch").Specific;
                    _Form.Items.Item("1").Visible = true;

                    if (mtx.VisualRowCount > 0)
                    {
                        Utilities.Application.SBO_Application.StatusBar.SetText("Calculation complete. Automatically saving document...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);

                        // 3. Programmatically click the 'Add' or 'Update' button (UID 1)
                        // This triggers the standard SAP Add/Update event sequence
                        _Form.Items.Item("1").Click();

                        _Form.Items.Item("1").Visible = false;
                    }
                    else
                    {
                        Utilities.Application.SBO_Application.StatusBar.SetText("Calculation finished, but no rows were found. Document not saved.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    }
                }
            }
            catch (Exception ex) { Utilities.Application.SBO_Application.MessageBox("Calc Error: " + ex.Message); }
            finally { _Form.Freeze(false); if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs); }
        }

        private void ProcessMonthlySegments(SAPbobsCOM.GeneralDataCollection oProvisions, string loanCode, DateTime monStart, DateTime monEnd, string monDocNum, double startingTotal)
        {
            SAPbobsCOM.Recordset rsSlabs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            double loanMonthlyTotal = startingTotal;

            try
            {
                string sqlSlabs = $@"SELECT T1.""U_FromDate"", T1.""U_ToDate"", T1.""U_InterestPer"", T0.""U_BalanceAmountC""
                            FROM ""@LOANMASTERDETAILSH"" T0 INNER JOIN ""@LOANMASTERDETAILSC1"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                            WHERE T0.""U_LoanCode"" = '{loanCode}' 
                            AND T1.""U_FromDate"" <= '{monEnd:yyyyMMdd}' AND T1.""U_ToDate"" >= '{monStart:yyyyMMdd}' AND T1.""U_Status"" = 'O'
                            AND T0.""U_BalanceAmountC"" > 0 ";
                rsSlabs.DoQuery(sqlSlabs);

                while (!rsSlabs.EoF)
                {
                    double rate = Convert.ToDouble(rsSlabs.Fields.Item("U_InterestPer").Value);
                    double defBal = Convert.ToDouble(rsSlabs.Fields.Item("U_BalanceAmountC").Value);
                    DateTime slabF = (DateTime)rsSlabs.Fields.Item("U_FromDate").Value;
                    DateTime slabT = (DateTime)rsSlabs.Fields.Item("U_ToDate").Value;

                    DateTime currentStart = slabF < monStart ? monStart : slabF;
                    DateTime segmentLimit = slabT > monEnd ? monEnd : slabT;

                    while (currentStart <= segmentLimit)
                    {


                        DateTime currentEnd = GetNextSegmentEnd(loanCode, currentStart, segmentLimit);

                        string rowRemark = "";
                        if (currentEnd < segmentLimit)
                        {
                            // If the segment was cut short before reaching the end of the interest slab/month
                            rowRemark = "Due to Principal Repayment";
                        }
                        else
                        {
                            // If the segment reached the end of the slab or month boundary
                            rowRemark = "Due to Interest Period/Rate Change";
                        }

                        double principal = GetPrincipalForDate(loanCode, currentStart, defBal);
                        int days = (currentEnd - currentStart).Days + 1;

                        double interest = Math.Round((principal * rate / 100.0) * (days / 365.0), 2);

                        // Add Detail Row
                        var newRow = oProvisions.Add();
                        newRow.SetProperty("U_FromDate", currentStart);
                        newRow.SetProperty("U_ToDate", currentEnd);
                        newRow.SetProperty("U_InterestPer", rate.ToString());
                        newRow.SetProperty("U_PrincipalAmtC", principal.ToString());
                        newRow.SetProperty("U_InterestAmtC", interest.ToString());
                        newRow.SetProperty("U_NoDays", days.ToString());
                        newRow.SetProperty("U_IntrDateBook", monEnd);
                        newRow.SetProperty("U_DIntrDateBook", monEnd);
                        newRow.SetProperty("U_NIntrDateBook", GetNepaliDate(monEnd));
                        newRow.SetProperty("U_Remarks", rowRemark);
                        newRow.SetProperty("U_MonDocNum", monDocNum);

                        loanMonthlyTotal += interest;
                        currentStart = currentEnd.AddDays(1);
                    }
                    rsSlabs.MoveNext();
                }

                // ADD SUMMARY ROW (Happens immediately after details)
                if (loanMonthlyTotal > 0)
                {
                    string nepDate = GetNepaliDate(monEnd);
                    string nepYM = nepDate.Length >= 7 ? nepDate.Substring(0, 7) : nepDate;

                    var summaryRow = oProvisions.Add();
                    summaryRow.SetProperty("U_InterestPer", $"Monthly Total Interest for (BS) {nepYM}");
                    summaryRow.SetProperty("U_DIntrDateBook", monEnd);
                    summaryRow.SetProperty("U_MonDocNum", monDocNum);
                    summaryRow.SetProperty("U_Remarks", Math.Round(loanMonthlyTotal, 2).ToString());
                    summaryRow.SetProperty("U_InterestAmtC", "0"); // Usually 0 so it doesn't double-count in reports
                }
            }
            finally { System.Runtime.InteropServices.Marshal.ReleaseComObject(rsSlabs); }
        }

        private ProvisionRowData ExtractRowData(SAPbobsCOM.GeneralData row)
        {
            return new ProvisionRowData
            {
                FromDate = row.GetProperty("U_FromDate"),
                ToDate = row.GetProperty("U_ToDate"),
                InterestPer = row.GetProperty("U_InterestPer")?.ToString(),
                Principal = row.GetProperty("U_PrincipalAmtC")?.ToString(),
                InterestAmt = row.GetProperty("U_InterestAmtC")?.ToString(),
                NoDays = row.GetProperty("U_NoDays")?.ToString(),
                IntrDateBook = row.GetProperty("U_IntrDateBook"),
                DIntrDateBook = row.GetProperty("U_DIntrDateBook"),
                NIntrDateBook = row.GetProperty("U_NIntrDateBook")?.ToString(),
                Remarks = row.GetProperty("U_Remarks")?.ToString(),
                MonDocNum = row.GetProperty("U_MonDocNum")?.ToString(),
                JENo = row.GetProperty("U_JENo")?.ToString()
            };
        }

        private void InsertRowData(SAPbobsCOM.GeneralData row, ProvisionRowData data)
        {
            row.SetProperty("U_FromDate", data.FromDate);
            row.SetProperty("U_ToDate", data.ToDate);
            row.SetProperty("U_InterestPer", data.InterestPer);
            row.SetProperty("U_PrincipalAmtC", data.Principal);
            row.SetProperty("U_InterestAmtC", data.InterestAmt);
            row.SetProperty("U_NoDays", data.NoDays);
            row.SetProperty("U_IntrDateBook", data.IntrDateBook);
            row.SetProperty("U_DIntrDateBook", data.DIntrDateBook);
            row.SetProperty("U_NIntrDateBook", data.NIntrDateBook);
            row.SetProperty("U_Remarks", data.Remarks);
            row.SetProperty("U_MonDocNum", data.MonDocNum);
            row.SetProperty("U_JENo", data.JENo);
        }

        private void PopulateFiscalYears(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item("cmbFisY").Specific;
            SAPbobsCOM.Recordset oRec = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                // Clear existing values
                while (oCombo.ValidValues.Count > 0)
                    oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);

                string sql = "SELECT DISTINCT \"Category\" FROM OFPR ORDER BY \"Category\" DESC";
                oRec.DoQuery(sql);

                while (!oRec.EoF)
                {
                    oCombo.ValidValues.Add(oRec.Fields.Item("Category").Value.ToString(), oRec.Fields.Item("Category").Value.ToString());
                    oRec.MoveNext();
                }
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.MessageBox(ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec);
            }
        }

        private void PopulatePostingPeriods(SAPbouiCOM.Form oForm, string selectedYear)
        {
            SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item("cmbPosP").Specific;
            SAPbobsCOM.Recordset oRec = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                while (oCombo.ValidValues.Count > 0)
                    oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // Fetch periods belonging to the selected year
                string sql = $"SELECT \"Code\", \"Name\" FROM OFPR WHERE \"Category\" = '{selectedYear}' ORDER BY \"AbsEntry\"";
                oRec.DoQuery(sql);

                while (!oRec.EoF)
                {
                    oCombo.ValidValues.Add(oRec.Fields.Item("Code").Value.ToString(), oRec.Fields.Item("Name").Value.ToString());
                    oRec.MoveNext();
                }
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.MessageBox(ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec);
            }
        }

        private void SetDatesFromPeriod(SAPbouiCOM.Form oForm, string periodAbsEntry)
        {
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                // Fetch Start and End dates for the selected Posting Period
                string sql = $"SELECT \"F_RefDate\", \"T_RefDate\" FROM OFPR WHERE \"Code\" = '{periodAbsEntry}'";
                rs.DoQuery(sql);

                if (!rs.EoF)
                {
                    DateTime sDate = (DateTime)rs.Fields.Item("F_RefDate").Value;
                    DateTime eDate = (DateTime)rs.Fields.Item("T_RefDate").Value;

                    // Set values to DataSources (Standard yyyyMMdd format for SAP EditText)
                    _MasterDBDataSource.SetValue("U_StartDate", 0, sDate.ToString("yyyyMMdd"));
                    _MasterDBDataSource.SetValue("U_EndDate", 0, eDate.ToString("yyyyMMdd"));

                    // Force the UI to refresh from the DataSource
                    ((SAPbouiCOM.EditText)oForm.Items.Item("txtSDate").Specific).Value = sDate.ToString("yyyyMMdd");
                    ((SAPbouiCOM.EditText)oForm.Items.Item("txtEDate").Specific).Value = eDate.ToString("yyyyMMdd");
                }
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.StatusBar.SetText("Error setting period dates: " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }


        private bool ValidatePreviousMonthBooked()
        {
            SAPbobsCOM.Recordset rs = null;

            try
            {
                string selectedStartDateStr =
                    _MasterDBDataSource.GetValue("U_StartDate", 0).Trim();

                string selectedEndDateStr =
                    _MasterDBDataSource.GetValue("U_EndDate", 0).Trim();

                if (string.IsNullOrEmpty(selectedStartDateStr) ||
                    string.IsNullOrEmpty(selectedEndDateStr))
                {
                    return false;
                }

                DateTime selectedStartDate =
                    DateTime.ParseExact(selectedStartDateStr, "yyyyMMdd", null);

                DateTime selectedEndDate =
                    DateTime.ParseExact(selectedEndDateStr, "yyyyMMdd", null);

                rs = (SAPbobsCOM.Recordset)
                    Utilities.Application.Company.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.BoRecordset);

                /*
                 * Get all loans for the currently selected month.
                 */
                string sqlLoans = $@"
            SELECT DISTINCT
                T1.""U_LoanCode""
            FROM ""@LOANMASTERDETAILSC2"" T0
            INNER JOIN ""@LOANMASTERDETAILSH""  T1 ON T0.""DocEntry"" = T1.""DocEntry""
            WHERE T0.""U_DIntrDateBook"" = '{selectedEndDate:yyyyMMdd}'
              AND IFNULL(T1.""U_LoanCode"", '') <> ''
            ORDER BY T1.""U_LoanCode""";

                rs.DoQuery(sqlLoans);

                while (!rs.EoF)
                {
                    string loanCode =
                        rs.Fields.Item("U_LoanCode").Value.ToString().Trim();

                    /*
                     * Find the immediately previous booking month
                     * for this SAME loan.
                     */
                    string sqlPrevious = $@"
                SELECT
                    MAX(T0.""U_DIntrDateBook"") AS ""PreviousBookDate""
                FROM ""@LOANMASTERDETAILSC2"" T0
                INNER JOIN ""@LOANMASTERDETAILSH""  T1 ON T0.""DocEntry"" = T1.""DocEntry""
                WHERE T1.""U_LoanCode"" = '{loanCode.Replace("'", "''")}'
                  AND T0.""U_DIntrDateBook"" < '{selectedEndDate:yyyyMMdd}'
                  AND T0.""U_FromDate"" IS NOT NULL";

                    SAPbobsCOM.Recordset rsPrevious =
                        (SAPbobsCOM.Recordset)
                        Utilities.Application.Company.GetBusinessObject(
                            SAPbobsCOM.BoObjectTypes.BoRecordset);

                    rsPrevious.DoQuery(sqlPrevious);

                    if (!rsPrevious.EoF &&
                        rsPrevious.Fields.Item("PreviousBookDate").Value != null)
                    {
                        DateTime previousBookDate =
                            (DateTime)rsPrevious.Fields.Item("PreviousBookDate").Value;

                        /*
                         * Now check whether there are any interest detail
                         * rows for that previous month which do NOT have JE.
                         */
                        string sqlUnbooked = $@"
                    SELECT COUNT(*) AS ""Cnt""
                    FROM ""@LOANMASTERDETAILSC2"" T0
                    INNER JOIN ""@LOANMASTERDETAILSH""  T1 ON T0.""DocEntry"" = T1.""DocEntry""
                    WHERE T1.""U_LoanCode"" =
                          '{loanCode.Replace("'", "''")}'
                      AND T0.""U_DIntrDateBook"" = '{previousBookDate:yyyyMMdd}'
                      AND T0.""U_FromDate"" IS NOT NULL
                      AND IFNULL(T0.""U_InterestAmtC"", 0) > 0
                      AND IFNULL(T0.""U_JENo"", '') = ''";

                        SAPbobsCOM.Recordset rsUnbooked =
                            (SAPbobsCOM.Recordset)
                            Utilities.Application.Company.GetBusinessObject(
                                SAPbobsCOM.BoObjectTypes.BoRecordset);

                        rsUnbooked.DoQuery(sqlUnbooked);

                        int unbookedCount =
                            Convert.ToInt32(rsUnbooked.Fields.Item("Cnt").Value);

                        if (unbookedCount > 0)
                        {
                            string previousNepaliDate =
                                GetNepaliDate(previousBookDate);

                            Utilities.Application.SBO_Application.MessageBox(
                                $"Cannot book interest for Loan {loanCode}.\n\n" +
                                $"Previous month's interest ({previousNepaliDate.Substring(0, 7)}) " +
                                $"is still not booked.\n\n" +
                                $"Please book the previous month's interest first.");

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(rsUnbooked);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(rsPrevious);

                            return false;
                        }

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(rsUnbooked);
                    }

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rsPrevious);

                    rs.MoveNext();
                }

                return true;
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.MessageBox(
                    "Previous month validation error: " + ex.Message);

                return false;
            }
            finally
            {
                if (rs != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }

        private class ProvisionRowData
        {
            public object FromDate;
            public object ToDate;
            public string InterestPer;
            public string Principal;
            public string InterestAmt;
            public string NoDays;
            public object IntrDateBook;
            public object DIntrDateBook;
            public string NIntrDateBook;
            public string Remarks;
            public string MonDocNum;
            public string JENo;
        }

    }
    

}
