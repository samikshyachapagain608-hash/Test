using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BankFacility
{
    class ClsLoanDisbursement : Base
    {
        #region CONSTRUCTOR AND DISTRUCTOR
        public ClsLoanDisbursement() : base()
        {

        }

        ~ClsLoanDisbursement()
        {

        }
        #endregion

        #region VARIABLE DECLARATION

        SAPbouiCOM.ChooseFromListEvent _SysCFLEvent = null;
        public SAPbouiCOM.DBDataSource _MasterDBDataSource = null;

        private string tempBankGL = "";
        private string tempLoanGL = "";
        private double tempAmount = 0;
        private string tempCostCenter = "";
        private string tempRemarks = "";
        private string tempRefNo = "";
        private string tempJournalR = "";
        private string tempLoanCode = "";
        private string tempLoanEndDate ;
        private string tempLoanStartDate;
        private string tempInterest = "";
        private string endDtStr = "";
        //private string tempBankGL = "";
        private SAPbouiCOM.BoFormMode _lastMode;
        private string tempDocEntry = "";
        private string tempBaseInterest = "";
        private string tempPremInterest = "";
        private string tempIncPay = "";
        private string tempNoDays = "";
        private string tempBalAmt = "";
        private string tempDocCode = "";
        private bool _isProcessing = false;
        string cashFlowId = "";
        string query1 = null;
        SAPbobsCOM.Recordset _RS = null, _RS1 = null;
        string startDate;
        string lType = "", fCode = "";

        private string facStartDate = "";
        private string facEndDate = "";
        #endregion

        #region Form Data Event
        public override void FormData_Event(ref BusinessObjectInfo BusinessObjectInfo, ref bool oBubbleEvent)
        {
            if (BusinessObjectInfo.BeforeAction == false && BusinessObjectInfo.ActionSuccess == true)
            {
                if (BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_LOAD)
                {
                    SAPbouiCOM.Form oForm = null;

                    LockFieldsAfterLoad();
                   
                }
            }
        }
        #endregion

        #region Form Default
        internal void FormDefault()
        {
            _Form.DataBrowser.BrowseBy = "txtDocE";
            _Form.Settings.Enabled = true;

            _MasterDBDataSource = _Form.DataSources.DBDataSources.Item("@LOANDISBURSEMENTH");
            if (_Form.Mode == BoFormMode.fm_ADD_MODE)
            {
                _MasterDBDataSource.SetValue("U_BookedBy", 0, Utilities.Application.Company.UserName);

                _MasterDBDataSource.SetValue("U_BaseInterest", 0, "0.0");
                _MasterDBDataSource.SetValue("U_PremInterest", 0, "0.0");
                _MasterDBDataSource.SetValue("U_Interest", 0, "0.0");

                SetNextDocNum();
                LockFieldsAfterLoad();
            }
            else if (_Form.Mode == BoFormMode.fm_FIND_MODE)
            {
                PrepareFormForFind();
            }
        }
        #endregion

        public override void Menu_Event(ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            if (pVal.BeforeAction == false)
            {
                if (pVal.MenuUID == "1281") // Find Mode
                {
                    PrepareFormForFind();
                }
                if (pVal.MenuUID == "1282") // Add Mode
                {
                    SetNextDocNum();
                    LockFieldsAfterLoad();
                }
            }
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
                        if (pVal.ItemUID == "1" && (_Form.Mode == BoFormMode.fm_ADD_MODE || _Form.Mode == BoFormMode.fm_UPDATE_MODE))
                        {
                            string existingPay = _MasterDBDataSource.GetValue("U_IncPayNo", 0).Trim();
                            startDate = _MasterDBDataSource.GetValue("U_LoanStartDate", 0).Trim();
                            _lastMode = _Form.Mode;
                            lType = _MasterDBDataSource.GetValue("U_LoanType", 0).Trim();
                            fCode = _MasterDBDataSource.GetValue("U_BankFacCode", 0).Trim();

                            if (!ValidateForm())
                            {
                                BubbleEvent = false; // Stop the "Add" or "Update" action
                                return;
                            }

                            if (!string.IsNullOrEmpty(existingPay) && existingPay != "0")
                            {
                                return;
                            }

                            //if (!string.IsNullOrEmpty(existingPay) && existingPay != "0")
                            //{
                            //    tempAmount = 0;
                            //}
                            //else
                            //{

                            tempDocCode = _MasterDBDataSource.GetValue("Code", 0).Trim();

                            if (!ValidateForm()) { BubbleEvent = false; return; }

                                try
                                {
                                    // Capture the current record ID (DocEntry for Document-type UDO)
                                    tempDocEntry = _MasterDBDataSource.GetValue("DocEntry", 0).Trim();
                                    tempLoanEndDate = _MasterDBDataSource.GetValue("U_LoanEndDate", 0).Trim();
                                    tempLoanStartDate = _MasterDBDataSource.GetValue("U_LoanStartDate", 0).Trim();
                                    tempInterest = _MasterDBDataSource.GetValue("U_Interest", 0).Trim();
                                    tempBankGL = _MasterDBDataSource.GetValue("U_BankGL", 0).Trim();
                                    tempLoanGL = _MasterDBDataSource.GetValue("U_LoanAC", 0).Trim();
                                    tempAmount = Convert.ToDouble(_MasterDBDataSource.GetValue("U_LoanAmountC", 0).Trim());
                                    tempCostCenter = _MasterDBDataSource.GetValue("U_CostCenterCode", 0).Trim();
                                    tempRemarks = _MasterDBDataSource.GetValue("U_Remarks", 0).Trim();
                                    tempRefNo = _MasterDBDataSource.GetValue("U_RefNo", 0).Trim();
                                    tempJournalR = _MasterDBDataSource.GetValue("U_JournalRemarks", 0).Trim();
                                    tempLoanCode = _MasterDBDataSource.GetValue("U_LoanCode", 0).Trim();
                                    tempBaseInterest = _MasterDBDataSource.GetValue("U_BaseInterest", 0).Trim();
                                    tempPremInterest = _MasterDBDataSource.GetValue("U_PremInterest", 0).Trim();
                                    tempNoDays = _MasterDBDataSource.GetValue("U_LoanPeriod", 0).Trim();
                                    tempBalAmt = _MasterDBDataSource.GetValue("U_LoanAmountC", 0).Trim();
                            }
                                catch { }
                            //}
                            if (!string.IsNullOrEmpty(tempLoanGL) && tempAmount > 0)
                            {
                                int response = Utilities.Application.SBO_Application.MessageBox(
                                $"Are you sure you want to disburse Loan {tempLoanCode} for amount {tempAmount:N2}? This will create a Journal Entry/Incoming Payment.",
                                1, "Yes", "No");

                                if (response != 1) // User clicked 'No' or closed the window
                                {
                                    BubbleEvent = false; 
                                    return;
                                }

                                string incomingPay =  CreateIncomingPaymentFromStashedData();
                                if (string.IsNullOrEmpty(incomingPay))
                                {
                                    // PAYMENT FAILED: 
                                    // The error message is already shown inside CreateIncomingPaymentDI.
                                    // By setting BubbleEvent to false, SAP cancels the "Add" or "Update" of your UDO.
                                    BubbleEvent = false;
                                    return;
                                }
                                else
                                {
                                    _MasterDBDataSource.SetValue("U_IncPayNo", 0, incomingPay);
                                }
                            }


                        }

                        if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "Item_91")
                        {
                            string docEntryToFind = _MasterDBDataSource.GetValue("U_FDocEntry", 0).Trim();

                            if (!string.IsNullOrEmpty(docEntryToFind))
                            {
                                BubbleEvent = false; // Stop SAP default action

                                try
                                {
                                    // Use a local 'Base' variable, do NOT overwrite 'this._Object'
                                    //Base targetLogic = new ClsBankFacility();

                                    ClsBankFacility objBankFact = new ClsBankFacility();
                                    objBankFact.IsViewOnly = true;

                                    Base targetLogic = objBankFact;

                                    Utilities.LoadForm(ref targetLogic, Constants.Forms.Bank_Facility);

                                    SAPbouiCOM.Form targetForm = Utilities.Application.SBO_Application.Forms.ActiveForm;

                                    // Simple validation to prevent the Null Reference
                                    if (targetForm == null) return;

                                    targetForm.Freeze(true);
                                    try
                                    {
                                        // Switch to FIND mode
                                        targetForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

                                        // Access the ID field. Ensure "txtDocE" is the UID in the Bank Facility XML
                                        SAPbouiCOM.Item oDocEntryItem = targetForm.Items.Item("txtDocE");
                                        oDocEntryItem.Enabled = true; // Must be enabled to set value in Find mode

                                        SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oDocEntryItem.Specific;
                                        oEdit.Value = docEntryToFind;

                                        // Click the Find Button (UID "1")
                                        targetForm.Items.Item("1").Click();

                                        // NOW close the original form
                                        _Form.Close();
                                    }
                                    finally
                                    {
                                        targetForm.Freeze(false);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Utilities.Application.SBO_Application.StatusBar.SetText("Error opening Facility: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                }
                            }
                        }


                        break;

                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        if (pVal.ItemUID == "txtLoanC")
                        {
                            ApplyLoanCodeFilter(FormUID);
                        }
                        if (pVal.ItemUID == "txtCFlw")
                        {
                            ApplyCashFlowFilter();
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
                        if (pVal.ItemUID == "1")
                        {
                                //if (_lastMode == BoFormMode.fm_ADD_MODE)
                                //{
                                    // Process using our stashed list
                                    //CreateCostCentersFromStash(_stashedRows);
                                    //if (!string.IsNullOrEmpty(tempLoanGL) && tempAmount > 0)
                                    //{
                                    //    CreateIncomingPaymentFromStashedData();
                                    //}
                                    //if (pVal.ActionSuccess)
                                    //{
                                    //    UpdateBankFacilityAvailedAmount(lType, fCode);
                                    //}
                                  // }



                        }
                       break;

                    case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                        if (pVal.ItemChanged && (pVal.ItemUID == "txtLnSD" || pVal.ItemUID == "txtLoanP"))
                        //if (pVal.ItemChanged && (pVal.ItemUID == "txtLnSD"))
                        {
                            CalculateEndDate();
                        }
                        if (pVal.ItemChanged && pVal.ItemUID == "txtLoanC")
                        {

                            FetchLoanMasterDetails(pVal.ItemUID);
                        }
                        if (pVal.ItemUID == "txtBInt" || pVal.ItemUID == "txtPInt")
                        {
                            CalculateTotalInterest();
                        }
                        break;

                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        _SysCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                        if (pVal.ItemUID == "txtLoanT")
                        {
                            string loanType = _SysCFLEvent.SelectedObjects.GetValue("Name", 0).ToString().Trim();
                            _MasterDBDataSource.SetValue("U_LoanType", 0, loanType);
                        }
                        if (pVal.ItemUID == "txtLoanC")
                        {
                            if (_SysCFLEvent.SelectedObjects != null)
                            {
                                if (pVal.ItemUID == "txtLoanC")
                                {
                                    string loanCode = _SysCFLEvent.SelectedObjects.GetValue("PrcCode", 0).ToString().Trim();
                                    _MasterDBDataSource.SetValue("U_LoanCode", 0, loanCode);

                                    // Trigger fetch immediately for CFL selection
                                    FetchLoanMasterDetails(pVal.ItemUID);
                                }
                            }
                        }
                        if(pVal.ItemUID == "txtCFlw")
                        {
                            string cashFlow = _SysCFLEvent.SelectedObjects.GetValue("CFWName", 0).ToString().Trim();
                             cashFlowId = _SysCFLEvent.SelectedObjects.GetValue("CFWId", 0).ToString().Trim();
                            _MasterDBDataSource.SetValue("U_CashFlow", 0, cashFlow);
                        }

                        break;
                }
            }
            #endregion
        }
        #endregion


        #region Calculate End Date
        private void CalculateEndDate()
        {
            try
            {
                _Form.Freeze(true);

                string startDateStr = _MasterDBDataSource.GetValue("U_LoanStartDate", 0).ToString().Trim();
                string daysStr = _MasterDBDataSource.GetValue("U_LoanPeriod", 0).ToString().Trim();

                if (!string.IsNullOrEmpty(startDateStr) && !string.IsNullOrEmpty(daysStr))
                {
                    DateTime startDate = DateTime.ParseExact(startDateStr, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);

                    if (double.TryParse(daysStr, out double days))
                    {
                        DateTime endDate = startDate.AddDays(days);

                        _MasterDBDataSource.SetValue("U_LoanEndDate", 0, endDate.ToString("yyyyMMdd"));
                        SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)_Form.Items.Item("txtLnED").Specific;
                        oEdit.Value = endDate.ToString("yyyyMMdd");
                    }
                }
            }
            catch (Exception ex)
            {
                // 
            }
            finally
            {
                _Form.Freeze(false);
            }
        }

        //private void CalculateEndDate()
        //{
        //    SAPbobsCOM.Recordset rs = null;
        //    try
        //    {
        //        _Form.Freeze(true);

        //        // 1. Get the English Start Date from DataSource
        //        string startDateStr = _MasterDBDataSource.GetValue("U_LoanStartDate", 0).ToString().Trim();
        //        if (string.IsNullOrEmpty(startDateStr)) return;

        //        DateTime engStartDate = DateTime.ParseExact(startDateStr, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);

        //        // 2. Convert English Start Date to Nepali Date
        //        string nepaliStartDate = GetNepaliDate(engStartDate); // Returns e.g., "2080/05/12"

        //        if (nepaliStartDate.Length >= 7)
        //        {
        //            // Extract "YYYY/MM" (e.g., "2080/05")
        //            string yrMonthPrefix = nepaliStartDate.Substring(0, 7);

        //            rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //            // 3. Query the calendar to find the English Date corresponding to the LAST Nepali day of that month
        //            string sql = $@"SELECT MAX(""U_EnglishDate"") as ""MaxEng""
        //                    FROM ""@NEPALICALENDER""
        //                    WHERE ""U_NepaliDate"" LIKE '{yrMonthPrefix}%'";

        //            rs.DoQuery(sql);

        //            if (!rs.EoF)
        //            {
        //                DateTime engLastDayOfMonth = (DateTime)rs.Fields.Item("MaxEng").Value;

        //                // 4. Update the End Date field in the DataSource and UI
        //                string finalEndDateStr = engLastDayOfMonth.ToString("yyyyMMdd");
        //                _MasterDBDataSource.SetValue("U_LoanEndDate", 0, finalEndDateStr);

        //                ((SAPbouiCOM.EditText)_Form.Items.Item("txtLnED").Specific).Value = finalEndDateStr;
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Utilities.Application.SBO_Application.StatusBar.SetText("Error calculating Nepali End Date: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //    finally
        //    {
        //        if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
        //        _Form.Freeze(false);
        //    }
        //}
        #endregion

        #region Get Nepali date
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
        #endregion

        #region Fetch the Loan Master Details on the basis of loan code
        private void FetchLoanMasterDetails(string itemUID)
        {
            if (_isProcessing) return;

            SAPbobsCOM.Recordset oRec = null, rsRem = null;
            SAPbobsCOM.Recordset oRec1 = null;
            try
            {
                _isProcessing = true;
                _Form.Freeze(true);
                oRec = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsRem = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRec1 = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                string selectedLoanCode = _MasterDBDataSource.GetValue("U_LoanCode", 0).ToString().Trim();
                string selectedLoanType = _MasterDBDataSource.GetValue("U_LoanType", 0).ToString().Trim();
                string selectedFacCode = "";

                if (string.IsNullOrEmpty(selectedLoanCode)) return;

                string sql = $@"SELECT T0.""U_BankCode"", T0.""U_BankName"", T1.""U_BankFacNo"", 
                                       T1.""U_CostCenterCode"", T1.""U_CostCenterName"", T1.""U_GLAccount"",T3.""U_BankGLCode"", T4.""DocEntry"",
                                       T4.""U_StartDate"" AS ""FacStart"", T4.""U_EndDate"" AS ""FacEnd"",
                             CASE WHEN '{selectedLoanType}' = 'TR Loan' THEN T3.""U_TrLoan"" WHEN '{selectedLoanType}' = 'Term Loan' THEN T3.""U_TermLoan""
                    WHEN '{selectedLoanType}' = 'Bridge Gap Loan' THEN T3.""U_BridgeLoan"" WHEN '{selectedLoanType}' = 'Demand Loan' THEN T3.""U_DemandLoan""
                    WHEN '{selectedLoanType}' = 'Over Draft Loan' THEN T3.""U_ODLoan"" WHEN '{selectedLoanType}' = 'Short Term Loan' THEN T3.""U_StlLoan""
                    END AS ""LoanGL"", T3.""U_BankGLCode"" ""BankGL""
                                FROM ""@LOANMASTERH"" T0
                                INNER JOIN ""@LOANMASTERC"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                                INNER JOIN ""@LOANGLAC"" T3 ON T0.""U_BankCode"" = T3.""Code""
                                INNER JOIN ""@BANKFACTH"" T4 ON T1.""U_BankFacNo"" = T4.""U_BnkFactNo""
                                WHERE T1.""U_LoanCode"" = '{selectedLoanCode}'";

                oRec.DoQuery(sql);

                selectedFacCode = oRec.Fields.Item("U_BanKFacNo").Value.ToString();
                object rawStart = oRec.Fields.Item("FacStart").Value;
                object rawEnd = oRec.Fields.Item("FacEnd").Value;

                facStartDate = (rawStart is DateTime dtStart) ? dtStart.ToString("yyyyMMdd") : rawStart.ToString();
                facEndDate = (rawEnd is DateTime dtEnd) ? dtEnd.ToString("yyyyMMdd") : rawEnd.ToString();

                string bankCode = oRec.Fields.Item("U_BankCode").Value.ToString();

                if (!string.IsNullOrEmpty(selectedFacCode))
                {
                    string sqlRem = $@"
            SELECT 
                CASE 
                    WHEN '{selectedLoanType}' = 'TR Loan' THEN IFNULL(""U_LCRemFact"", 0)
                    WHEN '{selectedLoanType}' = 'Short Term Loan' THEN IFNULL(""U_STLRemFact"", 0)
                    WHEN '{selectedLoanType}' = 'Over Draft Loan' THEN IFNULL(""U_ODRemFact"", 0)
                    WHEN '{selectedLoanType}' = 'Demand Loan' THEN IFNULL(""U_DMRemFact"", 0)
                    WHEN '{selectedLoanType}' = 'Bridge Gap Loan' THEN IFNULL(""U_BGRemFact"", 0)
                    WHEN '{selectedLoanType}' = 'Term Loan' THEN IFNULL(""U_TRMRemFact"", 0)
                    ELSE 0 
                END AS ""RemainingBalance""
            FROM ""@BANKFACTH"" 
            WHERE ""U_BnkFactNo"" = '{selectedFacCode.Replace("'", "''")}' and ""U_BnkCode"" = '{bankCode}'";

                    rsRem.DoQuery(sqlRem);

                    if (!rsRem.EoF)
                    {
                        double remBalance = Convert.ToDouble(rsRem.Fields.Item("RemainingBalance").Value);

                        _MasterDBDataSource.SetValue("U_RemAmountC", 0, remBalance.ToString());

                        ((SAPbouiCOM.EditText)_Form.Items.Item("txtRemA").Specific).Value = remBalance.ToString();

                    }
                }


                if (!oRec.EoF)
                {
                    _MasterDBDataSource.SetValue("U_LoanAC", 0, oRec.Fields.Item("LoanGL").Value.ToString());
                    _MasterDBDataSource.SetValue("U_PaymentMode", 0, "Bank Transfer");
                    _MasterDBDataSource.SetValue("U_BankCode", 0, oRec.Fields.Item("U_BankCode").Value.ToString());
                    _MasterDBDataSource.SetValue("U_BankName", 0, oRec.Fields.Item("U_BankName").Value.ToString());
                    _MasterDBDataSource.SetValue("U_BankFacCode", 0, oRec.Fields.Item("U_BanKFacNo").Value.ToString());
                    _MasterDBDataSource.SetValue("U_FDocEntry", 0, oRec.Fields.Item("DocEntry").Value.ToString());
                    _MasterDBDataSource.SetValue("U_CostCenterCode", 0, oRec.Fields.Item("U_CostCenterCode").Value.ToString());
                    _MasterDBDataSource.SetValue("U_CostCenterName", 0, oRec.Fields.Item("U_CostCenterName").Value.ToString());
                    _MasterDBDataSource.SetValue("U_BankGL", 0, oRec.Fields.Item("U_BankGLCode").Value.ToString());

                    _MasterDBDataSource.SetValue("U_BookedBy", 0, Utilities.Application.Company.UserName);

                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtBnkC").Specific).Value = oRec.Fields.Item("U_BankCode").Value.ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtBnkN").Specific).Value = oRec.Fields.Item("U_BankName").Value.ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtBFacC").Specific).Value = oRec.Fields.Item("U_BanKFacNo").Value.ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtCCC").Specific).Value = oRec.Fields.Item("U_CostCenterCode").Value.ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtCCN").Specific).Value = oRec.Fields.Item("U_CostCenterName").Value.ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtBGL").Specific).Value = oRec.Fields.Item("BankGL").Value.ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtBookB").Specific).Value = Utilities.Application.Company.UserName;
                }
                string sql1 = $@" SELECT T0.""TransId"",SUM(T0.""Credit"") ""TotalLoan"", T0.""RefDate"" FROM JDT1 T0 
                                   WHERE T0.""ProfitCode"" = '{selectedLoanCode.Replace("'", "''")}' OR T0.""OcrCode2"" = '{selectedLoanCode.Replace("'", "''")}'  OR T0.""OcrCode3"" = '{selectedLoanCode.Replace("'", "''")}'  
                                   OR T0.""OcrCode4"" = '{selectedLoanCode.Replace("'", "''")}'  OR T0.""OcrCode5"" = '{selectedLoanCode.Replace("'", "''")}' and T0.""Account"" = '{oRec.Fields.Item("U_GLAccount").Value.ToString()}'
                                    GROUP BY T0.""TransId"", T0.""RefDate"" ";
                oRec1.DoQuery(sql1);
                if (!oRec1.EoF)
                {
                    _MasterDBDataSource.SetValue("U_LoanAmountC", 0, oRec1.Fields.Item("TotalLoan").Value.ToString());
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtLnAmt").Specific).Value = oRec1.Fields.Item("TotalLoan").Value.ToString();

                }
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.StatusBar.SetText("Fetch Error: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (oRec != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec);
                _Form.Freeze(false);
                _isProcessing = false;
            }
        }
        #endregion

        #region Validate the Form
        private bool ValidateForm()
        {

            SAPbobsCOM.Recordset oRec = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                if (_Form.Mode == BoFormMode.fm_ADD_MODE)
                {
                    string currentCode = _MasterDBDataSource.GetValue("Code", 0).Trim();
                    if (string.IsNullOrEmpty(currentCode))
                    {
                        // Generate a unique ID (Timestamp is safest)
                        string newCode = DateTime.Now.ToString("yyyyMMddHHmmss");
                        _MasterDBDataSource.SetValue("Code", 0, newCode);
                        _MasterDBDataSource.SetValue("Name", 0, newCode);
                    }
                }


                string startDtStr = _MasterDBDataSource.GetValue("U_LoanStartDate", 0).Trim();
                if (string.IsNullOrEmpty(startDtStr))
                {
                    Utilities.Application.SBO_Application.StatusBar.SetText("Loan Start Date is mandatory.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return false;
                }

                if (!string.IsNullOrEmpty(facStartDate) && !string.IsNullOrEmpty(facEndDate))
                {
                    if (string.Compare(startDtStr, facStartDate) < 0 || string.Compare(startDtStr, facEndDate) > 0)
                    {
                        Utilities.Application.SBO_Application.StatusBar.SetText(
                            $"Error: Loan Start Date must be between Facility Start ({facStartDate}) and End ({facEndDate}).",
                            BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
                        return false;
                    }
                }

                string periodStr = _MasterDBDataSource.GetValue("U_LoanPeriod", 0).Trim();
                if (string.IsNullOrEmpty(periodStr) || periodStr == "0")
                {
                    Utilities.Application.SBO_Application.StatusBar.SetText("Loan Period is mandatory.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return false;
                }

                string endDtStr = _MasterDBDataSource.GetValue("U_LoanEndDate", 0).Trim();
                if (string.IsNullOrEmpty(endDtStr))
                {
                    Utilities.Application.SBO_Application.StatusBar.SetText("Loan End Date is mandatory.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return false;
                }

                string cashFlow = _MasterDBDataSource.GetValue("U_CashFlow", 0).Trim();
                if (string.IsNullOrEmpty(cashFlow))
                {
                    Utilities.Application.SBO_Application.StatusBar.SetText("Selection Cash Flow Type is mandatory.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return false;
                }

                // Validation: Loan Amount > 0
                // Change "U_LoanAmt" to your actual field name for the amount
                double loanAmt = 0;
                double remAmt = 0;
                double.TryParse(_MasterDBDataSource.GetValue("U_LoanAmountC", 0).Trim(), out loanAmt);
                double.TryParse(_MasterDBDataSource.GetValue("U_RemAmountC", 0).Trim(), out remAmt);
                if (loanAmt <= 0)
                {
                    Utilities.Application.SBO_Application.StatusBar.SetText("Loan Amount must be greater than 0.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return false;
                }

                if (loanAmt >= remAmt)
                {
                    Utilities.Application.SBO_Application.StatusBar.SetText($"Loan Amount ({loanAmt:N2}) cannot be greater than the Remaining Amount ({remAmt:N2}).", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return false;
                }



                // Validation: Bank G/L Account must be Active and Postable
                //string glAccount = _MasterDBDataSource.GetValue("U_LoanAC", 0).Trim();
                //if (string.IsNullOrEmpty(glAccount))
                //{
                //    Utilities.Application.SBO_Application.StatusBar.SetText("Bank G/L Account is missing.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                //    return false;
                //}

                //oRec.DoQuery($@"SELECT ""Postable"", ""Active"" FROM OACT WHERE ""AcctCode"" = '{glAccount}'");
                //if (oRec.EoF)
                //{
                //    Utilities.Application.SBO_Application.StatusBar.SetText("Invalid G/L Account.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                //    return false;
                //}
                //if (oRec.Fields.Item("Active").Value.ToString() != "Y" || oRec.Fields.Item("Postable").Value.ToString() != "Y")
                //{
                //    Utilities.Application.SBO_Application.StatusBar.SetText("Bank G/L Account must be Active and Postable.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                //    return false;
                //}

                // Validation: Loan Start Date >= Posting Period Start
                DateTime startDate = DateTime.ParseExact(startDtStr, "yyyyMMdd", null);

                // Fetch the start date of the Posting Period associated with the Start Date
                //oRec.DoQuery($@"SELECT ""F_RefDate"" FROM OFPR WHERE '{startDtStr}' BETWEEN ""F_RefDate"" AND ""T_RefDate""");
                //if (!oRec.EoF)
                //{
                //    DateTime periodStart = (DateTime)oRec.Fields.Item("F_RefDate").Value;
                //    if (startDate < periodStart)
                //    {
                //        Utilities.Application.SBO_Application.StatusBar.SetText("Loan Start Date cannot be earlier than the Posting Period Start Date.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                //        return false;
                //    }
                //}

                //  Validation: Loan End Date >= Start Date
                if (!string.IsNullOrEmpty(endDtStr))
                {
                    DateTime endDate = DateTime.ParseExact(endDtStr, "yyyyMMdd", null);
                    if (endDate < startDate)
                    {
                        Utilities.Application.SBO_Application.StatusBar.SetText("Loan End Date must be greater than or equal to Start Date.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        return false;
                    }
                }


                // for total interest validation
                string baseStr = _MasterDBDataSource.GetValue("U_BaseInterest", 0).Trim();
                string premStr = _MasterDBDataSource.GetValue("U_PremInterest", 0).Trim();

                if (string.IsNullOrEmpty(baseStr)) { _MasterDBDataSource.SetValue("U_BaseInterest", 0, "0"); baseStr = "0"; }
                if (string.IsNullOrEmpty(premStr)) { _MasterDBDataSource.SetValue("U_PremInterest", 0, "0"); premStr = "0"; }

                // Calculate sum for validation
                double baseVal = Convert.ToDouble(baseStr);
                double premVal = Convert.ToDouble(premStr);
                double totalVal = baseVal + premVal;

                // VALIDATION: Total Interest must be > 0
                if (totalVal <= 0)
                {
                    Utilities.Application.SBO_Application.StatusBar.SetText("Total Interest (Base + Premium) must be greater than 0%.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);

                    // Set focus back to the Base Interest field for the user
                    _Form.Items.Item("txtBInt").Click();
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.StatusBar.SetText("Validation Error: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                return false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec);
            }
        }
        #endregion

        #region Create Incoming Payment with Transaction and Cost Center Update
        private string CreateIncomingPaymentFromStashedData()
        {
            SAPbobsCOM.Payments oPayment = null;
            SAPbobsCOM.Recordset rsDim = null;
            bool transactionStarted = false;

            // Parse Dates
            DateTime postingDate = DateTime.Today;
            DateTime endDate = DateTime.Today;

            try
            {
                if (!string.IsNullOrEmpty(startDate))
                    postingDate = DateTime.ParseExact(startDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);

                if (!string.IsNullOrEmpty(tempLoanEndDate))
                    endDate = DateTime.ParseExact(tempLoanEndDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);

                // 1. START TRANSACTION
                if (!Utilities.Application.Company.InTransaction)
                {
                    Utilities.Application.Company.StartTransaction();
                    transactionStarted = true;
                }

                // 2. UPDATE COST CENTER DATES
                if (!string.IsNullOrEmpty(tempLoanCode))
                {
                    if (!UpdateCostCenterDates(tempLoanCode, postingDate, endDate))
                    {
                        string err = Utilities.Application.Company.GetLastErrorDescription();
                        if (Utilities.Application.Company.InTransaction)
                            Utilities.Application.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        Utilities.Application.SBO_Application.SetStatusBarMessage($"Failed to update Cost Center: {err}", BoMessageTime.bmt_Medium, true);
                        return "";
                    }
                }

                // 3. PREPARE INCOMING PAYMENT
                oPayment = (SAPbobsCOM.Payments)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.oIncomingPayments);
                rsDim = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);

                oPayment.DocType = BoRcptTypes.rAccount;
                oPayment.DocDate = postingDate;
                oPayment.BPLID = 1; // Adjust based on your environment
                oPayment.Remarks = tempRemarks;
                oPayment.JournalRemarks = tempJournalR;
                oPayment.CounterReference = tempRefNo;

                // Debit Side
                oPayment.TransferAccount = tempBankGL;
                oPayment.TransferSum = tempAmount;
                oPayment.TransferDate = postingDate;
                oPayment.TransferReference = tempRefNo;

                // Credit Side (Account Line)
                oPayment.AccountPayments.AccountCode = tempLoanGL;
                oPayment.AccountPayments.SumPaid = tempAmount;
                oPayment.AccountPayments.Decription = "Loan Disbursement Entry";

                // Assign Profit Center to the line based on Dimension
                if (!string.IsNullOrEmpty(tempLoanCode))
                {
                    string sql = $@"SELECT T0.""DimCode"" FROM OOCR T0 WHERE T0.""OcrCode"" = '{tempLoanCode.Replace("'", "''")}'";
                    rsDim.DoQuery(sql);
                    if (!rsDim.EoF)
                    {
                        int dimCode = Convert.ToInt32(rsDim.Fields.Item("DimCode").Value);
                        switch (dimCode)
                        {
                            case 1: oPayment.AccountPayments.ProfitCenter = tempLoanCode; break;
                            case 2: oPayment.AccountPayments.ProfitCenter2 = tempLoanCode; break;
                            case 3: oPayment.AccountPayments.ProfitCenter3 = tempLoanCode; break;
                            case 4: oPayment.AccountPayments.ProfitCenter4 = tempLoanCode; break;
                            case 5: oPayment.AccountPayments.ProfitCenter5 = tempLoanCode; break;
                        }
                    }
                }
                oPayment.AccountPayments.Add();

                oPayment.PrimaryFormItems.PaymentMeans = SAPbobsCOM.PaymentMeansTypeEnum.pmtBankTransfer;
                try
                {
                    oPayment.PrimaryFormItems.AmountLC = tempAmount;
                    oPayment.PrimaryFormItems.CashFlowLineItemID = Convert.ToInt32(cashFlowId);
                    oPayment.PrimaryFormItems.Add();
                }
                catch
                { //Handle if Cash Flow is not required for this GL
                }

                //System.Threading.Thread.Sleep(300);

                // POST PAYMENT
                int res = oPayment.Add();

                if (res != 0)
                {
                    string err = Utilities.Application.Company.GetLastErrorDescription();
                    // ROLLBACK Cost Center update because Payment failed
                    if (Utilities.Application.Company.InTransaction)
                        Utilities.Application.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                    Utilities.Application.SBO_Application.SetStatusBarMessage($"Payment Error: {err}", BoMessageTime.bmt_Medium, true);
                    return "";
                }
                else
                {
                    // 5. COMMIT TRANSACTION (Both CC Update and Payment succeeded)
                    string newPaymentKey = Utilities.Application.Company.GetNewObjectKey();
                    if (transactionStarted) Utilities.Application.Company.EndTransaction(BoWfTransOpt.wf_Commit);

                    UpdateBankFacilityAvailedAmount(lType, fCode, tempAmount);

                    // Update Loan Master Table (@LOANMASTERC) via SQL since JE is now created
                    UpdateLoanMasterTable(newPaymentKey);

                    double totalInt = 0, bInt = 0, pInt = 0;
                    double.TryParse(tempInterest, out totalInt);
                    double.TryParse(tempBaseInterest, out bInt);
                    double.TryParse(tempPremInterest, out pInt);

                    EnsureLoanMasterDetailsCreated(
                        tempLoanCode,
                        fCode,
                        tempAmount,
                        startDate,
                        tempLoanEndDate,
                        totalInt,
                        bInt,
                        pInt
                    );

                    return newPaymentKey;
                }
            }
            catch (Exception ex)
            {
                if (Utilities.Application.Company.InTransaction)
                    Utilities.Application.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                Utilities.Application.SBO_Application.SetStatusBarMessage("Payment Exception: " + ex.Message, BoMessageTime.bmt_Medium, true);
                return "";
            }
            finally
            {
                if (rsDim != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsDim);
                if (oPayment != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oPayment);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        #endregion


        #region Update Cost Center Dates 
        private bool UpdateCostCenterDates(string prcCode, DateTime start, DateTime end)
        {
            SAPbobsCOM.CompanyService oCompService = Utilities.Application.Company.GetCompanyService();
            SAPbobsCOM.ProfitCentersService oPCService = (SAPbobsCOM.ProfitCentersService)oCompService.GetBusinessService(SAPbobsCOM.ServiceTypes.ProfitCentersService);

            // Services use ProfitCenterParams to identify which record to act upon
            SAPbobsCOM.ProfitCenterParams oPCParams = (SAPbobsCOM.ProfitCenterParams)oPCService.GetDataInterface(SAPbobsCOM.ProfitCentersServiceDataInterfaces.pcsProfitCenterParams);
            oPCParams.CenterCode = prcCode;

            SAPbobsCOM.ProfitCenter oPC = null;

            try
            {
                // Get the existing Profit Center using the Params
                oPC = oPCService.GetProfitCenter(oPCParams);

                // Set the dates (Note the property names: EffectiveFrom and EffectiveTo)
                oPC.Effectivefrom = start;
                oPC.EffectiveTo = end;

                // Pass the modified object back to the service to update
                oPCService.UpdateProfitCenter(oPC);

                return true;
            }
            catch (Exception ex)
            {
                // If an error occurs here, the transaction in the calling method 
                // will catch it and perform the Rollback.
                Utilities.Application.SBO_Application.SetStatusBarMessage("Error updating Cost Center: " + ex.Message, BoMessageTime.bmt_Short, true);
                return false;
            }
            finally
            {
                if (oPC != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oPC);
                if (oPCParams != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oPCParams);
                if (oPCService != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oPCService);
            }
        }
        #endregion

        #region Update Loan Master Table with (dates, interest)
        private void UpdateLoanMasterTable(string newPaymentKey)
        {
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                // Formatting dates for SQL
                string sqlEndDate = string.IsNullOrEmpty(tempLoanEndDate) ? "NULL" : $"'{tempLoanEndDate}'";
                string sqlStartDate = string.IsNullOrEmpty(tempLoanStartDate) ? "NULL" : $"'{tempLoanStartDate}'";



                // Added U_StartDate = {sqlStartDate} to the update query
                string updateMasterSql = $@"
            UPDATE T0
            SET T0.""U_StartDate"" = {sqlStartDate},
                T0.""U_EndDate"" = {sqlEndDate},
                T0.""U_Interest"" = {tempInterest}, 
                T0.""U_BaseInterest"" = {tempBaseInterest}, 
                T0.""U_PremInterest"" = {tempPremInterest}, 
                T0.""U_JENO"" = '{newPaymentKey}', 
                T0.""U_NoDays"" = {tempNoDays},
                T0.""U_Status"" = 'Active', T0.""U_LoanAmountC"" = {tempAmount}
            FROM ""@LOANMASTERC"" T0
            WHERE T0.""U_LoanCode"" = '{tempLoanCode.Replace("'", "''")}'";

                rs.DoQuery(updateMasterSql);
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.StatusBar.SetText("Update Master Error: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }
        #endregion


        #region Calculate the total interest based on base and premium interest
        private void CalculateTotalInterest()
        {
            try
            {
                _Form.Freeze(true);
                double baseInt = 0;
                double premInt = 0;

                // Try to parse values from DataSource, default to 0 if blank or invalid
                double.TryParse(_MasterDBDataSource.GetValue("U_BaseInterest", 0).Trim(), out baseInt);
                double.TryParse(_MasterDBDataSource.GetValue("U_PremInterest", 0).Trim(), out premInt);

                double totalInt = baseInt + premInt;

                // Update the DataSource and the UI field (txtIntP)
                _MasterDBDataSource.SetValue("U_Interest", 0, totalInt.ToString());
                ((SAPbouiCOM.EditText)_Form.Items.Item("txtIntP").Specific).Value = totalInt.ToString();
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.StatusBar.SetText("Calc Error: " + ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
            }
        }
        #endregion


        #region Filter in Loan Code CFL
        private void ApplyLoanCodeFilter(string FormUID)
        {
            SAPbobsCOM.Recordset oRec = null;
            try
            {
                // Reset the CFL to ensure no old filters remain
                SAPbouiCOM.ChooseFromList oCFL = _Form.ChooseFromLists.Item("CFL_COC"); 
                oCFL.SetConditions(null);

                // Get the Loan Type selected on the form
                string selectedLoanType = _MasterDBDataSource.GetValue("U_LoanType", 0).Trim();

                // If no Loan Type is selected yet, don't show any loans
                if (string.IsNullOrEmpty(selectedLoanType))
                {
                    SetEmptyCFL(oCFL);
                    return;
                }

                oRec = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                // Query: Find all Loan Codes in the Master that match the Type AND have no JE Number
                // We do not filter by Bank Code here because you mentioned it's not on the form yet.
                string sql = $@"
            SELECT DISTINCT T0.""U_LoanCode"" 
            FROM ""@LOANMASTERC"" T0
            WHERE T0.""U_LoanType"" = '{selectedLoanType.Replace("'", "''")}' 
            AND (T0.""U_JENO"" IS NULL OR T0.""U_JENO"" = '')";

                oRec.DoQuery(sql);

                // Build Conditions based on the resulting PrcCodes
                SAPbouiCOM.Conditions oConds = new SAPbouiCOM.Conditions();
                int count = 0;

                while (!oRec.EoF)
                {
                    string validPrcCode = oRec.Fields.Item("U_LoanCode").Value.ToString().Trim();

                    if (!string.IsNullOrEmpty(validPrcCode))
                    {
                        if (count > 0) oConds.Item(count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;

                        SAPbouiCOM.Condition oCond = oConds.Add();
                        oCond.Alias = "PrcCode"; 
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCond.CondVal = validPrcCode;

                        count++;
                    }
                    oRec.MoveNext();
                }

                // Apply the conditions
                if (count == 0)
                {
                    // If no unprocessed loans exist for this type, show an empty list
                    SetEmptyCFL(oCFL);
                }
                else
                {
                    oCFL.SetConditions(oConds);
                }
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.StatusBar.SetText("CFL Filter Error: " + ex.Message);
            }
            finally
            {
                if (oRec != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec);
            }
        }
        #endregion

        #region Show entire cost centers if the filter doesnot match
        private void SetEmptyCFL(SAPbouiCOM.ChooseFromList oCFL)
        {
            SAPbouiCOM.Conditions oConds = new SAPbouiCOM.Conditions();
            SAPbouiCOM.Condition oCond = oConds.Add();
            oCond.Alias = "PrcCode";
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCond.CondVal = "-1"; // Force no results
            oCFL.SetConditions(oConds);
        }
        #endregion


        #region Fields freeze on the basis of mode
        private void LockFieldsAfterLoad()
        {
            if (Form.Mode == BoFormMode.fm_OK_MODE)
            {
                try
                {
                    _Form.Freeze(true);

                    _Form.Items.Item("txtLoanT").Enabled = false;
                    _Form.Items.Item("txtLoanC").Enabled = false;
                    _Form.Items.Item("txtLnSD").Enabled = false;
                    _Form.Items.Item("txtLoanP").Enabled = false;
                    _Form.Items.Item("txtBInt").Enabled = false;
                    _Form.Items.Item("txtPInt").Enabled = false;
                    _Form.Items.Item("txtLnAmt").Enabled = false;
                    _Form.Items.Item("txtRefNo").Enabled = false;
                    _Form.Items.Item("txtRmrk").Enabled = false;
                    _Form.Items.Item("txtJRmrk").Enabled = false;
                    _Form.Items.Item("txtCFlw").Enabled = false;

                }
                catch (Exception ex)
                {
                    Utilities.Application.SBO_Application.StatusBar.SetText("Error locking fields: " + ex.Message);
                }
                finally
                {
                    _Form.Freeze(false);
                }
            }

            if (Form.Mode == BoFormMode.fm_ADD_MODE)
            {
                try
                {
                    _Form.Freeze(true);

                    _Form.Items.Item("txtLoanT").Enabled = true;
                    _Form.Items.Item("txtLoanC").Enabled = true;
                    _Form.Items.Item("txtLnSD").Enabled = true;
                    _Form.Items.Item("txtLoanP").Enabled = true;
                    _Form.Items.Item("txtBInt").Enabled = true;
                    _Form.Items.Item("txtPInt").Enabled = true;
                    _Form.Items.Item("txtLnAmt").Enabled = true;
                    _Form.Items.Item("txtRefNo").Enabled = true;
                    _Form.Items.Item("txtRmrk").Enabled = true;
                    _Form.Items.Item("txtJRmrk").Enabled = true;
                    _Form.Items.Item("txtCFlw").Enabled = true;

                }
                catch (Exception ex)
                {
                    Utilities.Application.SBO_Application.StatusBar.SetText("Error locking fields: " + ex.Message);
                }
                finally
                {
                    _Form.Freeze(false);
                }
            }
        }
        #endregion


        //private void UpdateBankFacilityAvailedAmount(string loanType, string facilityCode)
        //{
        //    if (string.IsNullOrEmpty(loanType) || string.IsNullOrEmpty(facilityCode)) return;

        //    SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
        //    try
        //    {
        //        // 1. Calculate the total sum of disbursements for this Type and Facility
        //        string sumSql = $@"
        //    SELECT SUM(CAST(T0.""U_LoanAmount"" AS DECIMAL(21, 6))) as ""TotalSum""
        //    FROM ""@LOANDISBURSEMENTH"" T0 
        //    WHERE T0.""U_LoanType"" = '{loanType.Replace("'", "''")}' 
        //    AND T0.""U_BankFacCode"" = '{facilityCode.Replace("'", "''")}'";

        //        rs.DoQuery(sumSql);
        //        double totalAvailed = Convert.ToDouble(rs.Fields.Item("TotalSum").Value);

        //        // 2. Determine which column to update in @BANKFACTH
        //        string targetColumn = "";
        //        switch (loanType)
        //        {
        //            case "TR Loan": targetColumn = "U_LCFctAvld"; break; // Adjust names to match your UDFs
        //            case "Demand Loan": targetColumn = "U_DMFctAvld"; break;
        //            case "Short Term Loan": targetColumn = "U_STLFctAvld"; break;
        //            case "Over Draft Loan": targetColumn = "U_ODFctAvld"; break;
        //            case "Term Loan": targetColumn = "U_TRMFctAvld"; break;
        //            case "Bridge Gap Loan": targetColumn = "U_BGFctAvld"; break;
        //            default: return; // Type not mapped
        //        }

        //        if (!string.IsNullOrEmpty(targetColumn))
        //        {
        //            // 3. Update the Bank Facility Header
        //            string updateSql = $@"
        //        UPDATE ""@BANKFACTH"" 
        //        SET ""{targetColumn}"" = {totalAvailed} 
        //        WHERE ""U_BnkFactNo"" = '{facilityCode.Replace("'", "''")}'";

        //            rs.DoQuery(updateSql);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Utilities.Application.SBO_Application.StatusBar.SetText("Update Facility Error: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //    finally
        //    {
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
        //    }
        //}

        #region Update Amounts in Bank Facility 
        private void UpdateBankFacilityAvailedAmount(string loanType, string facilityCode, double currentAmount)
        {
            if (string.IsNullOrEmpty(loanType) || string.IsNullOrEmpty(facilityCode)) return;

            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                // 1. Calculate the total sum of disbursements for this Type and Facility
                string sumSql = $@"
            SELECT SUM(CAST(REPLACE(IFNULL(T0.""U_LoanAmountC"", '0'), ',', '') AS DECIMAL(21, 6))) as ""TotalSum""
            FROM ""@LOANDISBURSEMENTH"" T0 
            WHERE T0.""U_LoanType"" = '{loanType.Replace("'", "''")}' 
            AND T0.""U_BankFacCode"" = '{facilityCode.Replace("'", "''")}'";

                rs.DoQuery(sumSql);
                double existingSum = Convert.ToDouble(rs.Fields.Item("TotalSum").Value);
                double totalAvailed = existingSum + currentAmount;


                // 2. Map the columns
                string availCol = "";
                string totalCol = "";
                string remCol = "";

                switch (loanType)
                {
                    case "TR Loan": availCol = "U_LCFctAvld"; totalCol = "U_LCFactAmt"; remCol = "U_LCRemFact"; break;
                    case "Short Term Loan": availCol = "U_STLFctAvld"; totalCol = "U_STLFactAmt"; remCol = "U_STLRemFact"; break;
                    case "Over Draft Loan": availCol = "U_ODFctAvld"; totalCol = "U_ODFactAmt"; remCol = "U_ODRemFact"; break;
                    case "Demand Loan": availCol = "U_DMFctAvld"; totalCol = "U_DMFactAmt"; remCol = "U_DMRemFact"; break;
                    case "Bridge Gap Loan": availCol = "U_BGFctAvld"; totalCol = "U_BGFactAmt"; remCol = "U_BGRemFact"; break;
                    case "Term Loan": availCol = "U_TRMFctAvld"; totalCol = "U_TRMFactAmt"; remCol = "U_TRMRemFact"; break;
                }

                if (!string.IsNullOrEmpty(availCol))
                {
                    string updateSql = $@"
                UPDATE ""@BANKFACTH"" 
                SET ""{availCol}"" = {totalAvailed},
                    ""{remCol}"" = ""{totalCol}"" - {totalAvailed}
                WHERE ""U_BnkFactNo"" = '{facilityCode.Replace("'", "''")}'";


                    rs.DoQuery(updateSql);
                }
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.StatusBar.SetText("Update Facility Error: " + ex.Message);
            }
            finally { System.Runtime.InteropServices.Marshal.ReleaseComObject(rs); }
        }
        #endregion


        #region Document Number Increment 
        private void SetNextDocNum()
        {
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                // Query to get the max number. Cast to INT to ensure numeric sorting.
                string sql = "SELECT MAX(CAST(IFNULL(\"U_DocNum\", 0) AS INT)) FROM \"@LOANDISBURSEMENTH\"";
                rs.DoQuery(sql);

                int nextNum = 1;
                if (!rs.EoF)
                {
                    nextNum = Convert.ToInt32(rs.Fields.Item(0).Value) + 1;
                }

                _MasterDBDataSource.SetValue("U_DocNum", 0, nextNum.ToString());
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


        #region Form field enable for find mode
        private void PrepareFormForFind()
        {
            try
            {
                _Form.Freeze(true);

                // Loop through all items on the form
                for (int i = 0; i < _Form.Items.Count; i++)
                {
                    SAPbouiCOM.Item oItem = _Form.Items.Item(i);

                    // Disable all EditText and ComboBoxes
                    if (oItem.Type == BoFormItemTypes.it_EDIT || oItem.Type == BoFormItemTypes.it_COMBO_BOX)
                    {
                        // Only keep the DocNum field enabled
                        if (oItem.UniqueID == "txtDocNo")
                        {
                            oItem.Enabled = true;
                        }
                        else
                        {
                            oItem.Enabled = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.StatusBar.SetText("Find Mode Init Error: " + ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
            }
        }
        #endregion


        #region After loan Disbursement Create Entry in Loan MAster Details (if not created on double click)
        private void EnsureLoanMasterDetailsCreated(string loanCode, string facilityNo, double amount, string startDateStr, string endDateStr, double totalInt, double baseInt, double premInt)
        {
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.GeneralService oGeneralService;
            SAPbobsCOM.GeneralData oGeneralData;
            SAPbobsCOM.GeneralData oChild;
            SAPbobsCOM.GeneralDataCollection oChildren;
            SAPbobsCOM.CompanyService oCompService = Utilities.Application.Company.GetCompanyService();

            try
            {
                // Check if the record already exists
                string checkSql = $@"SELECT ""DocEntry"" FROM ""@LOANMASTERDETAILSH"" WHERE ""U_LoanCode"" = '{loanCode.Replace("'", "''")}'";
                rs.DoQuery(checkSql);

                if (!rs.EoF) return; // Already exists, do nothing

                // Get the end of the Nepali month for the Start Date
                string nepaliMonthEnd = "";
                string sqlNep = $@"SELECT MAX(T1.""U_EnglishDate"") as ""MaxEng"" 
                   FROM ""@NEPALICALENDER"" T1 
                   WHERE T1.""U_NepaliDate"" LIKE (SELECT SUBSTRING(T0.""U_NepaliDate"", 1, 7) || '%' 
                                                   FROM ""@NEPALICALENDER"" T0 
                                                   WHERE T0.""U_EnglishDate"" = '{startDateStr}')";
                rs.DoQuery(sqlNep);
                if (!rs.EoF)
                {
                    nepaliMonthEnd = ((DateTime)rs.Fields.Item("MaxEng").Value).ToString("yyyyMMdd");
                }

                // Create using General Service
                oGeneralService = oCompService.GetGeneralService("LOANMASTERDETAILS");
                oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);

                // Header Fields
                oGeneralData.SetProperty("U_LoanCode", loanCode);
                oGeneralData.SetProperty("U_FacilityNo", facilityNo);
                oGeneralData.SetProperty("U_LoanAmountC", amount.ToString());
                oGeneralData.SetProperty("U_PaidAmountC", 0.ToString());
                oGeneralData.SetProperty("U_BalanceAmountC", amount.ToString());
                oGeneralData.SetProperty("U_StartDate", DateTime.ParseExact(startDateStr, "yyyyMMdd", null));
                oGeneralData.SetProperty("U_EndDate", DateTime.ParseExact(endDateStr, "yyyyMMdd", null));

                // Add First Interest Slab (Child Table C1)
                oChildren = oGeneralData.Child("LOANMASTERDETAILSC1");
                oChild = oChildren.Add();

                DateTime dtStart = DateTime.ParseExact(startDateStr, "yyyyMMdd", null);
                oChild.SetProperty("U_FromDate", dtStart);

                if (!string.IsNullOrEmpty(nepaliMonthEnd))
                    oChild.SetProperty("U_ToDate", DateTime.ParseExact(nepaliMonthEnd, "yyyyMMdd", null));

                oChild.SetProperty("U_NFromDate",GetNepaliDate(dtStart));
                oChild.SetProperty("U_NToDate", GetNepaliDate(DateTime.ParseExact(nepaliMonthEnd, "yyyyMMdd", null)));

                oChild.SetProperty("U_InterestPer", totalInt.ToString());
                oChild.SetProperty("U_BaseInterestPer", baseInt.ToString());
                oChild.SetProperty("U_PremInterestPer", premInt.ToString());
                oChild.SetProperty("U_Status", "O");

                // Calculate Days
                if (!string.IsNullOrEmpty(nepaliMonthEnd))
                {
                    DateTime dtNepEnd = DateTime.ParseExact(nepaliMonthEnd, "yyyyMMdd", null);
                    int days = (dtNepEnd - dtStart).Days + 1;
                    oChild.SetProperty("U_NoDays", days);
                }

                // Add the record
                oGeneralService.Add(oGeneralData);
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.StatusBar.SetText("Auto-Master Creation Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }

        #endregion

        private void ApplyCashFlowFilter()
        {
            try
            {
                SAPbouiCOM.ChooseFromList oCFL = _Form.ChooseFromLists.Item("CFL_CFW");
                SAPbouiCOM.Conditions oConds = new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition oCond = oConds.Add();

                // Filter for Active items only
                oCond.Alias = "Postable";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "Y";

                oCFL.SetConditions(oConds);
            }
            catch (Exception ex)
            {
                // Handle if CFL is not found or already has conditions
            }
        }


    }
}
