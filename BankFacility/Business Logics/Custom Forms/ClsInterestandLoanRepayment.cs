using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BankFacility
{
    class ClsInterestandLoanRepayment : Base
    {
        #region Constructor and Distructor
        public ClsInterestandLoanRepayment() : base()
        {

        }

        ~ClsInterestandLoanRepayment() { }
        #endregion

        #region Variable Declaration
        SAPbouiCOM.ChooseFromListEvent _SysCFLEvent = null;
        public SAPbouiCOM.DBDataSource _MasterDBDataSource = null;
        SAPbobsCOM.Recordset _RS = null, _RS1 = null , _RS2 = null, _RS3 = null, rsDim = null;

        private Dictionary<string, string> _calendarCache = null;

        private Dictionary<string, string> _nepaliMonthNameCache = null;
        private Dictionary<string, DateTime> _nepaliMonthEndCache = null;

        string cashFlowId = "";
        #endregion

        #region Form Default
        internal void FormDefault()
        {
            Form.Freeze(true);
            _Form.DataBrowser.BrowseBy = "txtDocE";
            _Form.Settings.Enabled = true;
            _MasterDBDataSource = _Form.DataSources.DBDataSources.Item("@INTERESTLOANREPAY");

            //_Form.Items.Item("btnPost").Visible = false;

            //tabs and fields invisible 
            //_Form.Items.Item("Item_66").Visible = false;
            //_Form.Items.Item("Item_41").Visible = false;
            ////_Form.Items.Item("Item_42").Visible = false;
            ////_Form.Items.Item("txtAPIAC").Visible = false;
            //_Form.Items.Item("Item_44").Visible = false;
            //_Form.Items.Item("txtABnkAc").Visible = false;
            //_Form.Items.Item("Item_46").Visible = false;
            //_Form.Items.Item("txtAGLC").Visible = false;
            //_Form.Items.Item("Item_48").Visible = false;
            //_Form.Items.Item("txtAOPayN").Visible = false;
            //_Form.Items.Item("Item_50").Visible = false;
            //_Form.Items.Item("txtCrtBy").Visible = false;



            if (_Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                string currentDate = DateTime.Today.ToString("yyyyMMdd");
                _MasterDBDataSource.SetValue("U_PostingDate", 0, currentDate);
                _MasterDBDataSource.SetValue("U_PaymentMethod", 0, "Bank Transfer");

                try
                {
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtPostD").Specific).Value = currentDate;
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtPayM").Specific).Value = "Bank Transfer";
                }
                catch { }

                _Form.Items.Item("Item_3").Visible = false;
                _Form.Items.Item("Item_34").Visible = false;
                _Form.Items.Item("txtLAmt").Visible = false;
                _Form.Items.Item("Item_36").Visible = false;
                _Form.Items.Item("txtPRep").Visible = false;
                _Form.Items.Item("Item_38").Visible = false;
                _Form.Items.Item("txtPBal").Visible = false;
            }
            ManageFieldsEnabledState();

                Form.Freeze(false);
        }
        #endregion

        public override void FormData_Event(ref BusinessObjectInfo BusinessObjectInfo, ref bool oBubbleEvent)
        {
            if (BusinessObjectInfo.BeforeAction == false && BusinessObjectInfo.ActionSuccess == true)
            {
                if (BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_LOAD)
                {
                    HandleVisibility(FormUID);

                    ManageFieldsEnabledState();
                }

            }
        }
        #region ItemEvent
        public override void Item_Event(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            if (pVal.BeforeAction)
            {
                //if (pVal.EventType == BoEventTypes.et_CHOOSE_FROM_LIST && pVal.ItemUID == "txtLLoanC")
                //{
                //    ApplyCFLFilter(FormUID, "CFL_CC");
                //}

                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        if(pVal.ItemUID == "txtLLoanC")
                        {
                            ApplyLoanCodeFilter(FormUID);
                        }
                        if (pVal.ItemUID == "txtCashF")
                        {
                            ApplyCashFlowFilter();
                        }
                        break;

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        //if (pVal.ItemUID == "1" && pVal.BeforeAction == true && _Form.Mode == BoFormMode.fm_ADD_MODE)
                        //{
                        //    // Try creating the payment first
                        //    //if (!CreateOutgoingPayment())
                        //    //{
                        //    //    // If payment fails, stop the record from being added to the UDT
                        //    //    BubbleEvent = false;
                        //    //}

                        //    SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)_Form.Items.Item("txtPayT").Specific;
                        //    string selection = oCombo.Value.Trim();

                        //    // Basic Amount Validation (Required for both types)
                        //    string rawAmount = _MasterDBDataSource.GetValue("U_Amount", 0).Trim();
                        //    string refNo = _MasterDBDataSource.GetValue("U_RefNo", 0).Trim();
                        //    string journalRmk = _MasterDBDataSource.GetValue("U_JournalRmk", 0).Trim();
                        //    double amount = 0;
                        //    if (string.IsNullOrEmpty(rawAmount) || !double.TryParse(rawAmount, out amount) || amount <= 0)
                        //    {
                        //        Utilities.Application.SBO_Application.StatusBar.SetText("Please enter a valid Amount before proceeding.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        //        BubbleEvent = false;
                        //        return;
                        //    }
                        //    else if (string.IsNullOrEmpty(refNo))
                        //    {
                        //        Utilities.Application.SBO_Application.StatusBar.SetText("Please enter reference before proceeding.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        //        BubbleEvent = false;
                        //        return;
                        //    }
                        //    else if (string.IsNullOrEmpty(journalRmk))
                        //    {
                        //        Utilities.Application.SBO_Application.StatusBar.SetText("Please enter journal remark before proceeding.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        //        BubbleEvent = false;
                        //        return;
                        //    }

                        //    //Conditional Logic for Interest Payment
                        //    if (selection == "Interest Payment")
                        //    {
                        //        // Validation: Interest Bank Statement (txtIBnkS) must hold a value
                        //        string bankStmt = _MasterDBDataSource.GetValue("U_InterestBankS", 0).Trim();
                        //        if (string.IsNullOrEmpty(bankStmt))
                        //        {
                        //            Utilities.Application.SBO_Application.StatusBar.SetText("Interest Bank Statement is required for Interest Payments.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        //            _Form.Items.Item("txtIBnkS").Click();
                        //            BubbleEvent = false;
                        //            return;
                        //        }

                        //        // Create Journal Entry First
                        //        // We check if Difference is not zero. If it's 0, we might skip JE creation based on your logic
                        //        double diff = 0;
                        //        double.TryParse(_MasterDBDataSource.GetValue("U_Difference", 0).Trim(), out diff);

                        //        if (diff != 0)
                        //        {
                        //            if (!CreateJournalEntry())
                        //            {
                        //                BubbleEvent = false;
                        //                return;
                        //            }
                        //        }

                        //        // Create Outgoing Payment
                        //        if (!CreateOutgoingPayment())
                        //        {
                        //            BubbleEvent = false;
                        //            return;
                        //        }
                        //    }
                        //    // 4. Conditional Logic for Principal Repayment
                        //    else if (selection == "Principal Repayment")
                        //    {
                        //        // Create ONLY Outgoing Payment
                        //        if (!CreateOutgoingPayment())
                        //        {
                        //            BubbleEvent = false;
                        //            return;
                        //        }
                        //    }
                        //}

                        if (pVal.ItemUID == "1" && pVal.BeforeAction == true && _Form.Mode == BoFormMode.fm_ADD_MODE)
                        {
                            SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)_Form.Items.Item("txtPayT").Specific;
                            string selection = oCombo.Value.Trim();

                            string loanCode = _MasterDBDataSource.GetValue("U_LoanCode", 0).Trim();

                            // Basic Validations
                            string rawAmount = _MasterDBDataSource.GetValue("U_AmountC", 0).Trim();
                            string refNo = _MasterDBDataSource.GetValue("U_RefNo", 0).Trim();
                            string journalRmk = _MasterDBDataSource.GetValue("U_JournalRmk", 0).Trim();
                            double amount = 0;

                            string postDateStr = _MasterDBDataSource.GetValue("U_PostingDate", 0).Trim();
                            string todayStr = DateTime.Today.ToString("yyyyMMdd");

                            if (string.IsNullOrEmpty(postDateStr))
                            {
                                Utilities.Application.SBO_Application.StatusBar.SetText("Posting Date is mandatory.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                _Form.Items.Item("txtPostD").Click(); // Set focus to the field
                                BubbleEvent = false;
                                return;
                            }

                            if (string.Compare(postDateStr, todayStr) > 0)
                            {
                                Utilities.Application.SBO_Application.StatusBar.SetText($"Posting Date ({postDateStr}) cannot be greater than Current Date ({todayStr}).", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                _Form.Items.Item("txtPostD").Click();
                                BubbleEvent = false; return;
                            }

                            if (string.IsNullOrEmpty(rawAmount) || !double.TryParse(rawAmount, out amount) || amount <= 0)
                            {
                                Utilities.Application.SBO_Application.StatusBar.SetText("Please enter a valid Amount.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false; return;
                            }
                            if (string.IsNullOrEmpty(refNo) || string.IsNullOrEmpty(journalRmk))
                            {
                                Utilities.Application.SBO_Application.StatusBar.SetText("Ref No and Journal Remarks are required.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false; return;
                            }

                            SAPbobsCOM.Recordset rsCheck = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                            try
                            {
                                string sql = $@"
                                SELECT MAX(T1.""U_IntrDateBook"") as ""MaxBookDate""
                                FROM ""@LOANMASTERDETAILSH"" T0 
                                INNER JOIN ""@LOANMASTERDETAILSC2"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                                WHERE T0.""U_LoanCode"" = '{loanCode.Replace("'", "''")}' 
                                AND T1.""U_JENo"" IS NOT NULL";

                                rsCheck.DoQuery(sql);

                                if (!rsCheck.EoF)
                                {
                                    object val = rsCheck.Fields.Item("MaxBookDate").Value;

                                    // Use Convert.IsDBNull to safely check for NULL from SQL
                                    if (val != null && !Convert.IsDBNull(val))
                                    {
                                        DateTime maxBookDate = Convert.ToDateTime(val);
                                        string maxBookDateStr = maxBookDate.ToString("yyyyMMdd");

                                        if (string.Compare(postDateStr, maxBookDateStr) < 0)
                                        {
                                            Utilities.Application.SBO_Application.StatusBar.SetText(
                                                $"Error: For Interest Payment, date cannot be earlier than last Booking Date ({maxBookDateStr}).",
                                                BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);

                                            BubbleEvent = false;
                                            return;
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Utilities.Application.SBO_Application.StatusBar.SetText("Date Validation Query Error: " + ex.Message);
                                BubbleEvent = false; return;
                            }
                            finally
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(rsCheck);
                            }

                            // --- TRANSACTION START ---
                            if (Utilities.Application.Company.InTransaction) Utilities.Application.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                            Utilities.Application.Company.StartTransaction();

                            try
                            {
                                string newJEKey = "";
                                string newPayKey = "";
                                if (selection == "Interest Payment")
                                {
                                    // Validation for Bank Statement
                                    string bankStmt = _MasterDBDataSource.GetValue("U_InterestBankSC", 0).Trim();
                                    if (string.IsNullOrEmpty(bankStmt))
                                    {
                                        throw new Exception("Interest Bank Statement is required for Interest Payments.");
                                    }

                                    // Create Journal Entry (only if difference != 0)
                                    double diff = 0;
                                    double.TryParse(_MasterDBDataSource.GetValue("U_DifferenceC", 0).Trim(), out diff);

                                    if (diff != 0)
                                    {
                                        newJEKey = CreateJournalEntry(); 
                                    }
                                    newPayKey = CreateOutgoingPayment();

                                    string transId = "";
                                    if (!string.IsNullOrEmpty(newPayKey))
                                    {
                                        // Retrieve the TransId (Journal Entry ID) associated with this Outgoing Payment
                                        transId = GetTransIdFromPayment(newPayKey);

                                        // Now you can use transId as needed
                                        Utilities.Application.SBO_Application.StatusBar.SetText("Payment Created. JE TransID: " + transId, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                                    }

                                    // UpdateLoanMasterWithDI(loanCode, selection, amount, newPayKey);
                                    UpdateLoanMasterWithDI(loanCode, selection, amount,transId);
                                }
                                else if (selection == "Principal Repayment")
                                {
                                    newPayKey = CreateOutgoingPayment();

                                    string transId = "";
                                    if (!string.IsNullOrEmpty(newPayKey))
                                    {
                                        // Retrieve the TransId (Journal Entry ID) associated with this Outgoing Payment
                                        transId = GetTransIdFromPayment(newPayKey);

                                        // Now you can use transId as needed
                                        Utilities.Application.SBO_Application.StatusBar.SetText("Payment Created. JE TransID: " + transId, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                                    }

                                    UpdateLoanMasterWithDI(loanCode, selection, amount, transId);
                                }

                                // If we reached here, both (or the single one) succeeded
                                if (Utilities.Application.Company.InTransaction)
                                    Utilities.Application.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                                if (selection == "Interest Payment")
                                {
                                    if (!string.IsNullOrEmpty(newJEKey))
                                    {
                                        _MasterDBDataSource.SetValue("U_JENo", 0, newJEKey);
                                        try { ((SAPbouiCOM.EditText)_Form.Items.Item("txtJERef").Specific).Value = newJEKey; } catch { }
                                    }
                                }
                                if (!string.IsNullOrEmpty(newPayKey))
                                {
                                    if (selection == "Interest Payment")
                                        _MasterDBDataSource.SetValue("U_OutgoingPay", 0, newPayKey);
                                    else
                                        _MasterDBDataSource.SetValue("U_AOutgoingPay", 0, newPayKey);
                                }
                            }
                            catch (Exception ex)
                            {
                                if (Utilities.Application.Company.InTransaction)
                                    Utilities.Application.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                                Utilities.Application.SBO_Application.MessageBox("Transaction Failed: " + ex.Message);
                                BubbleEvent = false;
                                return;
                            }
                        }

                        break;
                }

            }

            else
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        //if (pVal.ItemUID == "btnPost")
                        //{
                        //    if (CreateJournalEntry())
                        //    {
                        //        // Hide button after successful posting
                        //        _Form.Items.Item("btnPost").Visible = false;
                        //    }
                        //}
                        
                        break;


                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:

                        _SysCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;

                        // IMPORTANT: Only process AFTER the selection is made (BeforeAction = false)
                        if (pVal.BeforeAction == false)
                        {
                            SAPbouiCOM.DataTable oDT = _SysCFLEvent.SelectedObjects;
                            if (oDT == null) return; // Exit if user cancelled the selection

                            if (pVal.ItemUID == "txtLoanT")
                            {
                                ClearFieldsAfter("LOAN_TYPE");
                                
                                string selectedLoanType = Convert.ToString(oDT.GetValue("Name", 0)).Trim();
                                _MasterDBDataSource.SetValue("U_LoanType", 0, selectedLoanType);
                            }

                            if (pVal.ItemUID == "txtBankC")
                            {
                                ClearFieldsAfter("BANK_CODE");

                                // Get the Bank Code from the current CFL selection
                                string bankCode = Convert.ToString(oDT.GetValue("U_BankCode", 0)).Trim();
                                _MasterDBDataSource.SetValue("U_BankCode", 0, bankCode);

                                // FETCH the Loan Type that was previously saved in the DataSource
                                string savedLoanType = _MasterDBDataSource.GetValue("U_LoanType", 0).Trim();

                                SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)_Form.Items.Item("txtPayT").Specific;
                                string selection = oCombo.Value.Trim();

                                //if(selection == "Principal Repayment")
                                //{
                                //    FetchLoanMasterDetails1(bankCode, savedLoanType);
                                //}

                            }

                            if (pVal.ItemUID == "txtLLoanC")
                            {
                                // Get Loan Code from selection
                                string loanCode = Convert.ToString(oDT.GetValue("PrcCode", 0)).Trim();
                                _MasterDBDataSource.SetValue("U_LoanCode", 0, loanCode);

                                // Fetch other details based on this Loan Code
                                string bankCode = _MasterDBDataSource.GetValue("U_BankCode", 0).Trim();
                                string loanType = _MasterDBDataSource.GetValue("U_LoanType", 0).Trim();
                                FetchLoanMasterDetails(bankCode, loanType, loanCode);
                            }

                            if (pVal.ItemUID == "txtBnkGL")
                            {
                                string bankGL = Convert.ToString(oDT.GetValue("AcctCode", 0)).Trim();
                                string bankGLName = Convert.ToString(oDT.GetValue("AcctName", 0)).Trim();
                                _MasterDBDataSource.SetValue("U_BankGLAcc", 0, bankGL);
                                _MasterDBDataSource.SetValue("U_BankGLName", 0, bankGLName);
                            }
                            if (pVal.ItemUID == "txtABnkAc")
                            {
                                string bankGL = Convert.ToString(oDT.GetValue("AcctCode", 0)).Trim();
                                string bankGLName = Convert.ToString(oDT.GetValue("AcctName", 0)).Trim();
                                _MasterDBDataSource.SetValue("U_ABankGLAcc", 0, bankGL);
                                _MasterDBDataSource.SetValue("U_ABankGLName", 0, bankGLName);
                            }

                            if (pVal.ItemUID == "txtCashF")
                            {
                                string cashFlow = _SysCFLEvent.SelectedObjects.GetValue("CFWName", 0).ToString().Trim();
                                cashFlowId = _SysCFLEvent.SelectedObjects.GetValue("CFWId", 0).ToString().Trim();
                                _MasterDBDataSource.SetValue("U_CashFlow", 0, cashFlow);

                              
                            }
                        }

                        break;
                    case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                        if (pVal.ItemUID == "txtIBnkS")
                        {
                            CalculateDifference();
                        }
                        break;


                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        if (pVal.ItemUID == "txtPayT")
                        {
                            ClearFieldsAfter("PAYMENT_TYPE");

                            HandleVisibility(FormUID);
                        }
                        break;
                }
            }
         }
        #endregion

        private string GetTransIdFromPayment(string docEntry)
        {
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                // Query the Outgoing Payment table (OVPM) for the TransId
                // TransId is the field that links to the OJDT (Journal Entry) table
                string sql = $@"SELECT ""TransId"" FROM ""OVPM"" WHERE ""DocEntry"" = '{docEntry}'";

                rs.DoQuery(sql);

                if (!rs.EoF)
                {
                    return (rs.Fields.Item("TransId").Value.ToString());
                }
                return "";
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.StatusBar.SetText("Error retrieving TransId: " + ex.Message);
                return "";
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }

        #region CFL Filter in Loan Code
        private void ApplyCFLFilter(string FormUID, string cflId)
        {
            try
            {
                SAPbouiCOM.Form oForm = Utilities.Application.SBO_Application.Forms.Item(FormUID);
                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(cflId);

                // Clear existing conditions to prevent duplication
                SAPbouiCOM.Conditions oConds = new SAPbouiCOM.Conditions();

                string[] allowedGroups = { "STR", "BDG", "OD", "TRM", "TR", "DM" };

                for (int i = 0; i < allowedGroups.Length; i++)
                {
                    SAPbouiCOM.Condition oCond = oConds.Add();
                    oCond.Alias = "GrpCode"; // The field in OCRC (Cost Center Groups) or PRC1
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCond.CondVal = allowedGroups[i];

                    // Add an 'OR' relationship between conditions, except for the last one
                    if (i < allowedGroups.Length - 1)
                    {
                        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                    }
                }

                // Apply the conditions to the CFL object
                oCFL.SetConditions(oConds);
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.StatusBar.SetText("CFL Filter Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Fetch Details from Loan Master on the basis of Loan Code
        private void FetchLoanMasterDetails(string bankCode, string loanType, string loanCode)
        {
            SAPbobsCOM.Recordset oRec = null;
            try
            {
                _Form.Freeze(true);
                oRec = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                // SQL Query provided by you
                string sql = $@"
                    SELECT T0.""U_LoanType"", T0.""U_BankFacNo"", T0.""U_StartDate"", T0.""U_EndDate"", T0.""U_LoanCode"", T0.""U_LoanAmountC"",
                           T1.""U_BankName""
                    FROM ""@LOANMASTERC"" T0
                    INNER JOIN ""@LOANMASTERH"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                    WHERE T0.""U_LoanCode"" = '{loanCode.Replace("'", "''")}' ";

                //WHERE T0.""U_LoanCode"" = '{loanCode.Replace("'", "''")}

                oRec.DoQuery(sql);

                string query1 = $@"SELECT ""U_ProvisionAc"" ,""U_BankGLCode"",
                            CASE WHEN '{loanType}' = 'TR Loan' THEN ""U_TrLoan"" WHEN '{loanType}' = 'Term Loan' THEN ""U_TermLoan""
                    WHEN '{loanType}' = 'Bridge Gap Loan' THEN ""U_BridgeLoan"" WHEN '{loanType}' = 'Demand Loan' THEN ""U_DemandLoan""
                    WHEN '{loanType}' = 'Over Draft Loan' THEN ""U_ODLoan"" WHEN '{loanType}' = 'Short Term Loan' THEN ""U_StlLoan""
                    END AS ""LoanGL"", 
                    CASE WHEN '{loanType}' = 'TR Loan' THEN ""U_TRInterestAc"" WHEN '{loanType}' = 'Term Loan' THEN ""U_TermInterestAc""
                    WHEN '{loanType}' = 'Bridge Gap Loan' THEN ""U_BridgeInterestAc"" WHEN '{loanType}' = 'Demand Loan' THEN ""U_DemandInterestAc""
                    WHEN '{loanType}' = 'Over Draft Loan' THEN ""U_ODInterestAc"" WHEN '{loanType}' = 'Short Term Loan' THEN ""U_STLInterestAc""
                    END AS ""InterestAc""
                    FROM ""@LOANGLAC"" WHERE ""Code"" = '{bankCode.Replace("'", "''")}' ";
                Utilities.ExecuteSQL(ref _RS1, query1);

                if (!oRec.EoF)
                {
                    //_MasterDBDataSource.SetValue("U_LoanCode", 0, oRec.Fields.Item("U_LoanCode").Value.ToString());
                    _MasterDBDataSource.SetValue("U_LoanName", 0, oRec.Fields.Item("U_LoanType").Value.ToString());
                    _MasterDBDataSource.SetValue("U_FacilityCode", 0, oRec.Fields.Item("U_BankFacNo").Value.ToString());
                    _MasterDBDataSource.SetValue("U_BankName", 0, oRec.Fields.Item("U_BankName").Value.ToString());

                    // Dates must be formatted as yyyyMMdd for DBDataSource
                    DateTime startD = (DateTime)oRec.Fields.Item("U_StartDate").Value;
                    DateTime endD = (DateTime)oRec.Fields.Item("U_EndDate").Value;
                    _MasterDBDataSource.SetValue("U_StartDate", 0, startD.ToString("yyyyMMdd"));
                    _MasterDBDataSource.SetValue("U_EndDate", 0, endD.ToString("yyyyMMdd"));

                    string provAcc = _RS1.Fields.Item("U_ProvisionAc").Value.ToString();
                    _MasterDBDataSource.SetValue("U_ProvIntAcc", 0, _RS1.Fields.Item("U_ProvisionAc").Value.ToString());
                    _MasterDBDataSource.SetValue("U_IntAcc", 0, _RS1.Fields.Item("InterestAc").Value.ToString());
                    //_MasterDBDataSource.SetValue("U_BankGLAcc", 0, _RS1.Fields.Item("U_BankGLCode").Value.ToString());

                    //_MasterDBDataSource.SetValue("U_InterestBalance", 0, oRec.Fields.Item("U_LoanAmount").Value.ToString());

                    if (_MasterDBDataSource.GetValue("U_PaymentType", 0) == "Interest Payment")
                    {
                        string sqlLoanAmt1 = $@" SELECT
                    SUM(T0.""Debit"") AS ""TotalDebit"", SUM(T0.""Credit"") AS ""TotalCredit""
                FROM JDT1 T0
                INNER JOIN OJDT T1 ON T0.""TransId"" = T1.""TransId""
                WHERE
                    (T0.""ProfitCode"" = '{loanCode.Replace("'", "''")}' 
                     OR T0.""OcrCode2"" = '{loanCode.Replace("'", "''")}' 
                     OR T0.""OcrCode3"" = '{loanCode.Replace("'", "''")}' 
                     OR T0.""OcrCode4"" = '{loanCode.Replace("'", "''")}' 
                     OR T0.""OcrCode5"" = '{loanCode.Replace("'", "''")}')


                    AND T0.""Account"" = '{provAcc}'

                    AND IFNULL(T1.""StornoToTr"", 0) = 0

                    AND T1.""TransId"" NOT IN(
                        SELECT DISTINCT A.""StornoToTr""
                        FROM OJDT A
                        WHERE IFNULL(A.""StornoToTr"", 0) <> 0
                    )";   ///AND T0.""Debit"" > 0 after account
                        Utilities.ExecuteSQL(ref _RS3, sqlLoanAmt1);

                        double intBook = Convert.ToDouble(_RS3.Fields.Item("TotalCredit").Value);
                        double intpaid = Convert.ToDouble(_RS3.Fields.Item("TotalDebit").Value);
                        double intBalBook = intBook - intpaid;

                        _MasterDBDataSource.SetValue("U_TotalIntBookC", 0, intBook.ToString());
                        _MasterDBDataSource.SetValue("U_TotalIntPaidC", 0, intpaid.ToString());
                        _MasterDBDataSource.SetValue("U_InterestBalanceC", 0, intBalBook.ToString());

                    }


                        //For Repayment
                        string glCode = _RS1.Fields.Item("LoanGL").Value.ToString();
                    _MasterDBDataSource.SetValue("U_AProvIntAcc", 0, _RS1.Fields.Item("U_ProvisionAc").Value.ToString());
                    //_MasterDBDataSource.SetValue("U_ABankGLAcc", 0, _RS1.Fields.Item("U_BankGLCode").Value.ToString());
                    _MasterDBDataSource.SetValue("U_GLCode", 0, _RS1.Fields.Item("LoanGL").Value.ToString());
                    _MasterDBDataSource.SetValue("U_CreatedBy", 0, Utilities.Application.Company.UserName);

                    if (_MasterDBDataSource.GetValue("U_PaymentType", 0) == "Principal Repayment")
                    {
                        double loanAmount = Convert.ToDouble(oRec.Fields.Item("U_LoanAmountC").Value);
                        _MasterDBDataSource.SetValue("U_LoanAmountC", 0, loanAmount.ToString());

                        string sqlLoanAmt = $@" SELECT
                    SUM(T0.""Debit"") AS ""Total""
                FROM JDT1 T0
                INNER JOIN OJDT T1 ON T0.""TransId"" = T1.""TransId""
                WHERE
                    (T0.""ProfitCode"" = '{loanCode.Replace("'", "''")}' 
                     OR T0.""OcrCode2"" = '{loanCode.Replace("'", "''")}' 
                     OR T0.""OcrCode3"" = '{loanCode.Replace("'", "''")}' 
                     OR T0.""OcrCode4"" = '{loanCode.Replace("'", "''")}' 
                     OR T0.""OcrCode5"" = '{loanCode.Replace("'", "''")}')


                    AND T0.""Account"" = '{glCode}' 
                    AND T0.""Debit"" > 0

                    AND IFNULL(T1.""StornoToTr"", 0) = 0

                    AND T1.""TransId"" NOT IN(
                        SELECT DISTINCT A.""StornoToTr""
                        FROM OJDT A
                        WHERE IFNULL(A.""StornoToTr"", 0) <> 0
                    )";
                        Utilities.ExecuteSQL(ref _RS2, sqlLoanAmt);

                        double totalPrincipalR = Convert.ToDouble(_RS2.Fields.Item("Total").Value);

                        //Calculate the Principal Balance
                        double principalBalance = loanAmount - totalPrincipalR;

                        _MasterDBDataSource.SetValue("U_TotalPrincipalRC", 0, totalPrincipalR.ToString());
                        _MasterDBDataSource.SetValue("U_PrincipalBalanceC", 0, principalBalance.ToString());

                    }



                    // 2. Update the UI Fields (EditText)
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtLoanN").Specific).Value = oRec.Fields.Item("U_LoanType").Value.ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtFacC").Specific).Value = oRec.Fields.Item("U_BankFacNo").Value.ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtBnkN").Specific).Value = oRec.Fields.Item("U_BankName").Value.ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtStrtD").Specific).Value = startD.ToString("yyyyMMdd");
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtEndD").Specific).Value = endD.ToString("yyyyMMdd");
                    //((SAPbouiCOM.EditText)_Form.Items.Item("txtPIAcc").Specific).Value = oRec.Fields.Item("U_ProvisionAccount").Value.ToString();
                    //((SAPbouiCOM.EditText)_Form.Items.Item("txtBnkGL").Specific).Value = oRec.Fields.Item("U_BankGLAcc").Value.ToString();
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
            }
        }

        #endregion

        #region Visibility according to principal or Interest Selected in dropdown
        private void HandleVisibility(string FormUID)
        {
            try
            {
                _Form.Freeze(true);
                SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)_Form.Items.Item("txtPayT").Specific;
                string selection = oCombo.Value.Trim();

                // Interest Group UIDs
                string[] intGroup = { "txtTIBook", "Item_30", "txtTIPaid", "Item_32", "txtIntBal", "Item_28", "txtVar", "Item_9", "txtJERef", "Item_13", "txtIBnkS", "Item_15", "txtDiff", "Item_16",
                "Item_18", "txtBnkGL", "txtOPayNo", "txtBGLN", "txtJERmk", "Item_37", "Item_45"};
                // Principal Group UIDs
                string[] princGroup = { "Item_34", "txtLAmt", "Item_36", "txtPRep", "Item_38", "txtPBal", "txtABnkAc", "txtAGLN", "txtAOPayN", "Item_44", "Item_42", "Item_48"};

                if (selection == "Interest Payment") // Interest Payment
                {
                    SetGroupVisibility(intGroup, true, true);
                    // Hide Principal Group
                    SetGroupVisibility(princGroup, false, false);
                }
                else if (selection == "Principal Repayment") // Principal Repayment
                {
                    // Hide Interest Group
                    SetGroupVisibility(intGroup, false, false);
                    // Show Principal Group - Editable
                    SetGroupVisibility(princGroup, true, true);
                    //_Form.Items.Item("btnPost").Visible = false;
                }
                else
                {
                    // If nothing selected, hide both or show both as needed
                    SetGroupVisibility(intGroup, false, false);
                    SetGroupVisibility(princGroup, false, false);
                }
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.StatusBar.SetText(ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
            }
        }
        #endregion

        #region visibility according to dropdown
        private void SetGroupVisibility(string[] uids, bool visible, bool enabled)
        {
            foreach (string uid in uids)
            {
                try
                {
                    _Form.Items.Item(uid).Visible = visible;
                    // Labels (type 8) cannot be 'Enabled', so we check type
                    //if (_Form.Items.Item(uid).Type != BoFormItemTypes.it_STATIC)
                    //{
                    //    _Form.Items.Item(uid).Enabled = enabled;
                    //}
                }
                catch { /* Item might not exist in this pane */ }
            }
        }
        #endregion

        #region Create Journal Entry
        private string CreateJournalEntry()
        {
            SAPbobsCOM.JournalEntries oJE = null;
            try
            {
                string loanCode = _MasterDBDataSource.GetValue("U_LoanCode", 0).Trim();
                double diffAmount = 0;
                string valDiff = _MasterDBDataSource.GetValue("U_DifferenceC", 0).Trim();
                string remarks = _MasterDBDataSource.GetValue("U_JERemarks", 0).Trim();
                double.TryParse(valDiff, out diffAmount);

                // If the difference is 0, we don't need an adjustment entry
                //if (diffAmount == 0) return true;

                string provAcc = _MasterDBDataSource.GetValue("U_ProvIntAcc", 0).Trim();
                string intAcc = _MasterDBDataSource.GetValue("U_IntAcc", 0).Trim();

                if (string.IsNullOrEmpty(provAcc) || string.IsNullOrEmpty(intAcc))
                {
                    Utilities.Application.SBO_Application.MessageBox("G/L Accounts (Provision or Interest) are missing.");
                    //return false;
                }

                oJE = (SAPbobsCOM.JournalEntries)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                rsDim = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                // Determine Dimension Code once
                int dimCode = 0;
                if (!string.IsNullOrEmpty(loanCode))
                {
                    string sqlDim = $@"SELECT T0.""DimCode"" FROM OOCR T0 WHERE T0.""OcrCode"" = '{loanCode.Replace("'", "''")}'";
                    rsDim.DoQuery(sqlDim);
                    if (!rsDim.EoF)
                    {
                        dimCode = Convert.ToInt32(rsDim.Fields.Item("DimCode").Value);
                    }
                }


                // Use the Posting Date from the form if available, otherwise Today
                string postDateStr = _MasterDBDataSource.GetValue("U_PostingDate", 0).Trim();
                if (!string.IsNullOrEmpty(postDateStr))
                    oJE.ReferenceDate = DateTime.ParseExact(postDateStr, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                //else
                //    oJE.ReferenceDate = DateTime.Today;

                //oJE.Memo = "Interest Difference Adjustment";
                oJE.Memo = remarks.ToString();
                double absAmount = Math.Abs(diffAmount);

                oJE.Lines.AccountCode = provAcc;
                oJE.Lines.BPLID = 1;
                if (diffAmount > 0) oJE.Lines.Credit = absAmount; else oJE.Lines.Debit = absAmount;

                // Apply Dimension to Line 1
                ApplyDimensionToJELine(oJE, dimCode, loanCode);
                oJE.Lines.Add();

                // --- Create Line 2 (Interest Account) ---
                oJE.Lines.AccountCode = intAcc;
                oJE.Lines.BPLID = 1;
                if (diffAmount > 0) oJE.Lines.Debit = absAmount; else oJE.Lines.Credit = absAmount;

                // Apply Dimension to Line 2
                ApplyDimensionToJELine(oJE, dimCode, loanCode);
                oJE.Lines.Add();

                int result = oJE.Add();

                if (result == 0)
                {
                    return Utilities.Application.Company.GetNewObjectKey();
                    //string newEntry = Utilities.Application.Company.GetNewObjectKey();
                    //_MasterDBDataSource.SetValue("U_JENo", 0, newEntry);

                    //// UI Update
                    //try { ((SAPbouiCOM.EditText)_Form.Items.Item("txtJERef").Specific).Value = newEntry; } catch { }

                    //Utilities.Application.SBO_Application.StatusBar.SetText("Journal Entry " + newEntry + " created successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    //return true;
                }
                else
                {
                    //string error = Utilities.Application.Company.GetLastErrorDescription();
                    //Utilities.Application.SBO_Application.MessageBox("JE Error: " + error);
                    //return false;
                    throw new Exception("JE Error: " + Utilities.Application.Company.GetLastErrorDescription());
                }
            }
            //catch (Exception ex)
            //{
            //    //Utilities.Application.SBO_Application.StatusBar.SetText("JE Exception: " + ex.Message);
            //    //return false;
            //    string error = Utilities.Application.Company.GetLastErrorDescription();
            //    throw new Exception("Payment Error: " + error);
            //}
            finally
            {
                if (oJE != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oJE);
            }
        }
        #endregion

        #region Calculate Difference
        private void CalculateDifference()
        {
            try
            {
                _Form.Freeze(true);

                // Get values from DBDataSource (more reliable than reading text directly)
                double interestBal = 0;
                double bankStmt = 0;

                string val1 = _MasterDBDataSource.GetValue("U_InterestBalanceC", 0).Trim();
                string val2 = _MasterDBDataSource.GetValue("U_InterestBankSC", 0).Trim();

                double.TryParse(val1, out interestBal);
                double.TryParse(val2, out bankStmt);

                double difference = bankStmt - interestBal;

                // Set value to Difference field
                _MasterDBDataSource.SetValue("U_DifferenceC", 0, difference.ToString());

                // Update UI button visibility
                // Show button only if difference is not zero
                //if (difference != 0)
                //{
                //    _Form.Items.Item("btnPost").Visible = true;
                //}
                //else
                //{
                //    _Form.Items.Item("btnPost").Visible = false;
                //}
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

        #region Loan Code Filter according to Bank code and Loan Type
        private void ApplyLoanCodeFilter(string FormUID)
        {
            SAPbobsCOM.Recordset oRec = null;
            try
            {
                SAPbouiCOM.ChooseFromList oCFL = _Form.ChooseFromLists.Item("CFL_LoanC");
                oCFL.SetConditions(null); // Reset

                SAPbouiCOM.Conditions oConds = new SAPbouiCOM.Conditions();

                // Get the values from your DataSources
                string selectedLoanType = _MasterDBDataSource.GetValue("U_LoanType", 0).Trim();
                string selectedBankCode = _MasterDBDataSource.GetValue("U_BankCode", 0).Trim();

                if (string.IsNullOrEmpty(selectedLoanType) || string.IsNullOrEmpty(selectedBankCode))
                {
                    // if bank code and loan type is empty, set a dummy condition so the list is empty
                    SAPbouiCOM.Condition oCond = oConds.Add();
                    oCond.Alias = "PrcCode";
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCond.CondVal = "-1";
                    oCFL.SetConditions(oConds);
                    return;
                }

                // Execute SQL to find which Cost Centers (PrcCode) match the criteria
                oRec = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string sql = $@"
            SELECT T0.""PrcCode"" 
            FROM OPRC T0
            INNER JOIN ""@LOANMASTERC"" T1 ON T0.""PrcCode"" = T1.""U_LoanCode""
            INNER JOIN ""@LOANMASTERH"" T2 ON T1.""DocEntry"" = T2.""DocEntry""
            WHERE T1.""U_LoanType"" = '{selectedLoanType.Replace("'", "''")}' 
            AND T2.""U_BankCode"" = '{selectedBankCode.Replace("'", "''")}'";

                oRec.DoQuery(sql);

                // Loop through the results and build the CFL Conditions
                int count = 0;
                while (!oRec.EoF)
                {
                    if (count > 0) oConds.Item(count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;

                    SAPbouiCOM.Condition oCond = oConds.Add();
                    oCond.Alias = "PrcCode"; // Field in OPRC
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCond.CondVal = oRec.Fields.Item("PrcCode").Value.ToString();

                    count++;
                    oRec.MoveNext();
                }

                // case where no matching loans are found
                if (count == 0)
                {
                    SAPbouiCOM.Condition oCond = oConds.Add();
                    oCond.Alias = "PrcCode";
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCond.CondVal = "-1"; // Force empty list
                }

                oCFL.SetConditions(oConds);
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

        #region Create Outgoing Payment
        private string CreateOutgoingPayment()
        {
            SAPbobsCOM.Payments oPayment = null;
            SAPbobsCOM.Recordset rsDim = null;

            try
            {
                oPayment = (SAPbobsCOM.Payments)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
                rsDim = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                // 1. Get current selection from Dropdown
                SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)_Form.Items.Item("txtPayT").Specific;
                string selection = oCombo.Value.Trim();

                string bankGL = "";
                string targetGL = "";
                double amount = 0;
                string loanCode = _MasterDBDataSource.GetValue("U_LoanCode", 0).Trim();
                string refNo = "";
                string journalRmk = "";
                string postDateStr = _MasterDBDataSource.GetValue("U_PostingDate", 0).Trim();
                DateTime postingDate = DateTime.Today;
                if (!string.IsNullOrEmpty(postDateStr))
                {
                    postingDate = DateTime.ParseExact(postDateStr, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                }

                refNo = _MasterDBDataSource.GetValue("U_RefNo", 0).Trim();
                journalRmk = _MasterDBDataSource.GetValue("U_JournalRmk", 0).Trim();

                // 2. Set variables based on Selection
                if (selection == "Interest Payment")
                {
                    bankGL = _MasterDBDataSource.GetValue("U_BankGLAcc", 0).Trim();
                    targetGL = _MasterDBDataSource.GetValue("U_ProvIntAcc", 0).Trim();
                    double.TryParse(_MasterDBDataSource.GetValue("U_AmountC", 0).Trim(), out amount);
                }
                else if (selection == "Principal Repayment")
                {
                    bankGL = _MasterDBDataSource.GetValue("U_ABankGLAcc", 0).Trim();
                    targetGL = _MasterDBDataSource.GetValue("U_GLCode", 0).Trim();
                    double.TryParse(_MasterDBDataSource.GetValue("U_AmountC", 0).Trim(), out amount);
                }

                //if (amount <= 0 || string.IsNullOrEmpty(targetGL) || string.IsNullOrEmpty(bankGL))
                if (string.IsNullOrEmpty(targetGL) || string.IsNullOrEmpty(bankGL))
                {
                    Utilities.Application.SBO_Application.StatusBar.SetText("Invalid amount or missing G/L accounts.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    //return false;
                }

                //Header Information
                oPayment.DocType = SAPbobsCOM.BoRcptTypes.rAccount;
                //oPayment.DocDate = DateTime.Today;
                oPayment.DocDate = postingDate;
                oPayment.BPLID = 1;
                oPayment.Remarks = journalRmk;
                oPayment.JournalRemarks = journalRmk;
                oPayment.CounterReference = refNo;

                //Payment Means (Bank Transfer)
                oPayment.TransferAccount = bankGL;
                oPayment.TransferSum = amount;
                //oPayment.TransferSum = 500;
                //oPayment.TransferDate = DateTime.Today;
                oPayment.TransferDate = postingDate;
                oPayment.TransferReference = journalRmk;


                oPayment.PrimaryFormItems.PaymentMeans = SAPbobsCOM.PaymentMeansTypeEnum.pmtBankTransfer;
                oPayment.PrimaryFormItems.CashFlowLineItemID = Convert.ToInt32(cashFlowId);
                //oPayment.PrimaryFormItems.AmountLC = 500;
                oPayment.PrimaryFormItems.AmountLC = amount;
                oPayment.PrimaryFormItems.Add();

                // Account Line
                oPayment.AccountPayments.AccountCode = targetGL;
                //oPayment.AccountPayments.SumPaid = 500;
                oPayment.AccountPayments.SumPaid = amount;
                oPayment.AccountPayments.Decription = selection;

                if (!string.IsNullOrEmpty(loanCode))
                {
                    string sqlDim = $@"SELECT T0.""DimCode"" FROM OOCR T0 WHERE T0.""OcrCode"" = '{loanCode.Replace("'", "''")}'";
                    rsDim.DoQuery(sqlDim);

                    if (!rsDim.EoF)
                    {
                        int dimCode = Convert.ToInt32(rsDim.Fields.Item("DimCode").Value);
                        // Map to the correct Profit Center field in the Account Line
                        switch (dimCode)
                        {
                            case 1: oPayment.AccountPayments.ProfitCenter = loanCode; break;
                            case 2: oPayment.AccountPayments.ProfitCenter2 = loanCode; break;
                            case 3: oPayment.AccountPayments.ProfitCenter3 = loanCode; break;
                            case 4: oPayment.AccountPayments.ProfitCenter4 = loanCode; break;
                            case 5: oPayment.AccountPayments.ProfitCenter5 = loanCode; break;
                        }
                    }
                }

                oPayment.AccountPayments.Add();

                int result = oPayment.Add();

                if (result == 0)
                {
                    return Utilities.Application.Company.GetNewObjectKey();
                    //string newKey = Utilities.Application.Company.GetNewObjectKey();

                    //// Save the payment number
                    //if (selection == "Interest Payment")
                    //    _MasterDBDataSource.SetValue("U_OutgoingPay", 0, newKey);
                    //else
                    //    _MasterDBDataSource.SetValue("U_AOutgoingPay", 0, newKey);

                    //Utilities.Application.SBO_Application.StatusBar.SetText($"Outgoing Payment {newKey} created successfully.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                    //return true;
                }
                else
                {
                    //string error = Utilities.Application.Company.GetLastErrorDescription();
                    //Utilities.Application.SBO_Application.MessageBox($"Payment Error: {error}");
                    //return false;
                    throw new Exception("Payment Error: " + Utilities.Application.Company.GetLastErrorDescription());
                }
            }
            //catch (Exception ex)
            //{
            //    //Utilities.Application.SBO_Application.StatusBar.SetText($"Exception: {ex.Message}");
            //    //return false;
            //    string error = Utilities.Application.Company.GetLastErrorDescription();
            //    throw new Exception("Payment Error: " + error);
            //}
            finally
            {
                if (oPayment != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oPayment);
                if (rsDim != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsDim);
            }
        }

        #endregion

        #region Apply dimension to JE (Lines) Posting 
        private void ApplyDimensionToJELine(SAPbobsCOM.JournalEntries oJE, int dimCode, string loanCode)
        {
            if (dimCode <= 0 || string.IsNullOrEmpty(loanCode)) return;

            switch (dimCode)
            {
                case 1: oJE.Lines.CostingCode = loanCode; break;
                case 2: oJE.Lines.CostingCode2 = loanCode; break;
                case 3: oJE.Lines.CostingCode3 = loanCode; break;
                case 4: oJE.Lines.CostingCode4 = loanCode; break;
                case 5: oJE.Lines.CostingCode5 = loanCode; break;
            }
        }
        #endregion


        //private void UpdateLoanMasterWithDI(string loanCode, string paymentType, double amount, string jeKey)
        //{
        //    SAPbobsCOM.Recordset rs = null;
        //    try
        //    {
        //        // 1. Get current Header totals from the Master table
        //        rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //        string sql = $@"SELECT ""DocEntry"", ""U_BalanceAmount"", ""U_PaidAmount"" 
        //                FROM ""@LOANMASTERDETAILSH"" WHERE ""U_LoanCode"" = '{loanCode.Replace("'", "''")}'";
        //        rs.DoQuery(sql);

        //        if (rs.EoF) return;

        //        int masterDocEntry = Convert.ToInt32(rs.Fields.Item("DocEntry").Value);
        //        double runningBalance = Convert.ToDouble(rs.Fields.Item("U_BalanceAmount").Value);
        //        double totalPaid = Convert.ToDouble(rs.Fields.Item("U_PaidAmount").Value);

        //        // 2. Initialize GeneralService for the Loan Master UDO
        //        SAPbobsCOM.CompanyService oCompService = Utilities.Application.Company.GetCompanyService();
        //        SAPbobsCOM.GeneralService oGeneralService = oCompService.GetGeneralService("LOANMASTERDETAILS");
        //        SAPbobsCOM.GeneralDataParams oParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
        //        oParams.SetProperty("DocEntry", masterDocEntry);

        //        // 3. Get the Master Record object
        //        SAPbobsCOM.GeneralData oMasterData = oGeneralService.GetByParams(oParams);
        //        SAPbobsCOM.GeneralDataCollection oRepaymentHistory = oMasterData.Child("LOANMASTERDETAILSC");

        //        // 4. PERFORM CALCULATIONS (Exactly like the Master Form logic)
        //        double unpaidBefore = runningBalance;

        //        // Only Principal Repayments affect the running balance and header totals
        //        if (paymentType == "Principal Repayment")
        //        {
        //            runningBalance = unpaidBefore - amount;
        //            totalPaid += amount;

        //            // Update Header Properties (Passing doubles directly, no ToString needed)
        //            oMasterData.SetProperty("U_PaidAmount", totalPaid.ToString());
        //            oMasterData.SetProperty("U_BalanceAmount", runningBalance.ToString());
        //        }

        //        // 5. ADD THE ROW TO THE CHILD TABLE (@LOANMASTERDETAILSC)
        //        SAPbobsCOM.GeneralData oNewRow = oRepaymentHistory.Add();

        //        if (oRepaymentHistory.Count == 1 && string.IsNullOrEmpty(oRepaymentHistory.Item(0).GetProperty("U_JENO").ToString()))
        //        {
        //            oNewRow = oRepaymentHistory.Item(0);
        //        }
        //        else
        //        {
        //            oNewRow = oRepaymentHistory.Add();
        //        }

        //        // U_SrNo = current count (which includes the row we just added)
        //        oNewRow.SetProperty("U_SrNo", oRepaymentHistory.Count.ToString());
        //        oNewRow.SetProperty("U_PaymentDate", DateTime.Today);
        //        oNewRow.SetProperty("U_JENO", jeKey);

        //        // Numeric values passed as doubles
        //        oNewRow.SetProperty("U_UnPaidAmount", unpaidBefore.ToString());
        //        oNewRow.SetProperty("U_PaymentAmount", amount.ToString());
        //        oNewRow.SetProperty("U_BalanceAmount", runningBalance.ToString());

        //        // Nepali Date string
        //        oNewRow.SetProperty("U_NPaymentDate", GetNepaliDate(DateTime.Today));

        //        // 6. Final Update
        //        oGeneralService.Update(oMasterData);
        //    }
        //    catch (Exception ex)
        //    {
        //        // Caught by the transaction wrapper in Item_Event for rollback
        //        throw new Exception("Repayment Sync Error: " + ex.Message);
        //    }
        //    finally
        //    {
        //        if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
        //    }
        //}

        private void UpdateLoanMasterWithDI(string loanCode, string paymentType, double amount, string jeKey)
        {
            SAPbobsCOM.Recordset rs = null;
            try
            {
                // 1. Get current Header totals
                rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string sql = $@"SELECT ""DocEntry"", ""U_BalanceAmountC"", ""U_PaidAmountC"" 
                        FROM ""@LOANMASTERDETAILSH"" WHERE ""U_LoanCode"" = '{loanCode.Replace("'", "''")}'";
                rs.DoQuery(sql);

                if (rs.EoF) return;

                int masterDocEntry = Convert.ToInt32(rs.Fields.Item("DocEntry").Value);
                double runningBalance = Convert.ToDouble(rs.Fields.Item("U_BalanceAmountC").Value);
                double totalPaid = Convert.ToDouble(rs.Fields.Item("U_PaidAmountC").Value);

                // 2. Initialize GeneralService
                SAPbobsCOM.CompanyService oCompService = Utilities.Application.Company.GetCompanyService();
                SAPbobsCOM.GeneralService oGeneralService = oCompService.GetGeneralService("LOANMASTERDETAILS");
                SAPbobsCOM.GeneralDataParams oParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oParams.SetProperty("DocEntry", masterDocEntry);

                // 3. Get the Master Record
                SAPbobsCOM.GeneralData oMasterData = oGeneralService.GetByParams(oParams);
                SAPbobsCOM.GeneralDataCollection oRepaymentHistory = oMasterData.Child("LOANMASTERDETAILSC");

                // 4. Calculations
                double unpaidBefore = runningBalance;
                if (paymentType == "Principal Repayment")
                {
                    runningBalance = unpaidBefore - amount;
                    totalPaid += amount;

                    // Update Header (Using strings as per your requirement to match previous math)
                    oMasterData.SetProperty("U_PaidAmountC", totalPaid.ToString());
                    oMasterData.SetProperty("U_BalanceAmountC", runningBalance.ToString());
                }

                // 5. ADD/UPDATE ROW LOGIC (Corrected to prevent double-adding)
                SAPbobsCOM.GeneralData oNewRow;

                // Check if the first row is empty (ghost row)
                if (oRepaymentHistory.Count == 1 && string.IsNullOrEmpty(oRepaymentHistory.Item(0).GetProperty("U_JENO").ToString()))
                {
                    oNewRow = oRepaymentHistory.Item(0);
                }
                else
                {
                    // Add ONLY ONE new row here
                    oNewRow = oRepaymentHistory.Add();
                }

                //DateTime postingDate = DateTime.ParseExact("20260721", "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);


                // 6. Set Properties (Passing values to match your UDT precisely)
                // SrNo = total count of rows
                string postingDate = _MasterDBDataSource
                    .GetValue("U_PostingDate", 0)
                    .Trim();

                DateTime paymentDate = DateTime.ParseExact(
                    postingDate,
                    "yyyyMMdd",
                    System.Globalization.CultureInfo.InvariantCulture
                );
                oNewRow.SetProperty("U_SrNo", oRepaymentHistory.Count.ToString());
                //oNewRow.SetProperty("U_PaymentDate", DateTime.Today);
                oNewRow.SetProperty("U_PaymentDate", paymentDate);
                oNewRow.SetProperty("U_JENO", jeKey);
                oNewRow.SetProperty("U_UnPaidAmountC", unpaidBefore.ToString());
                oNewRow.SetProperty("U_PaymentAmountC", amount.ToString());
                oNewRow.SetProperty("U_BalanceAmountC", runningBalance.ToString());
                oNewRow.SetProperty("U_NPaymentDate", GetNepaliDate(DateTime.Today));

                // 7. Final Update
                oGeneralService.Update(oMasterData);
            }
            catch (Exception ex)
            {
                throw new Exception("Repayment Sync Error: " + ex.Message);
            }
            finally
            {
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }



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


        private void ClearFieldsAfter(string level)
        {
            if (_Form.Mode != BoFormMode.fm_ADD_MODE) return;

            _Form.Freeze(true);
            try
            {
                if (level == "PAYMENT_TYPE")
                {
                    _MasterDBDataSource.SetValue("U_LoanType", 0, "");
                    _MasterDBDataSource.SetValue("U_BankCode", 0, "");
                }
                else if (level == "LOAN_TYPE")
                {
                    _MasterDBDataSource.SetValue("U_BankCode", 0, "");
                }

                string[] fieldsToClear = {
            "U_LoanCode", "U_LoanName", "U_FacilityCode", "U_BankName", "U_GLCode",
            "U_StartDate", "U_EndDate", "U_InterestBalanceC", "U_InterestBankSC",
            "U_DifferenceC", "U_TotalIntBookC", "U_TotalIntPaidC", "U_AmountC",
            "U_RefNo", "U_JournalRmk", "U_LoanAmountC", "U_TotalPrincipalRC",
            "U_PrincipalBalanceC", "U_BankGLAcc", "U_BankGLName", "U_ProvIntAcc",
            "U_IntAcc", "U_ABankGLAcc", "U_ABankGLName", "U_CashFlow"
        };

                foreach (string field in fieldsToClear)
                {
                    _MasterDBDataSource.SetValue(field, 0, "");
                }

                ((SAPbouiCOM.EditText)_Form.Items.Item("txtLoanN").Specific).Value = "";
                ((SAPbouiCOM.EditText)_Form.Items.Item("txtBnkN").Specific).Value = "";

            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.StatusBar.SetText("Clear Error: " + ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
            }
        }

        private void ManageFieldsEnabledState()
        {
            try
            {
                _Form.Freeze(true);

                bool isOkMode = _Form.Mode == BoFormMode.fm_OK_MODE;
                bool isEditable = !isOkMode;

                string[] inputFields = {
            "txtPayT", "txtLoanT", "txtBankC", "txtLLoanC", "txtPostD",
            "txtAmt", "txtRefNo", "txtJRmk", "txtCashF", "txtIBnkS",
            "txtJERmk", "txtBnkGL", "txtABnkAc"
        };

                foreach (string uid in inputFields)
                {
                    try
                    {
                        _Form.Items.Item(uid).Enabled = isEditable;
                    }
                    catch { /* Item might be hidden or not on current pane */ }
                }
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.StatusBar.SetText("Lock Error: " + ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
            }
        }
    }
}
