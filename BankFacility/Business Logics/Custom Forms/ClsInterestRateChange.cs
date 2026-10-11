using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BankFacility
{
    class ClsInterestRateChange: Base
    {
        #region Constructor and Distructor
        public ClsInterestRateChange() : base()
        {

        }

        ~ClsInterestRateChange()
        {

        }
        #endregion


        #region VARIABLE DECLARATION
        SAPbouiCOM.ChooseFromListEvent _SysCFLEvent = null;
        public SAPbouiCOM.DBDataSource _MasterDBDataSource = null, _ChildDBDataSource = null;
        #endregion

        #region Item Event
        public override void Item_Event(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            //before action
            if (pVal.BeforeAction)
            {
            }
            //after action
            else
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:

                        _SysCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;

                        SAPbouiCOM.DataTable oDT = _SysCFLEvent.SelectedObjects;
                        if (oDT == null) return;

                        if (pVal.ItemUID == "txtLoanT")
                        {
                            string selectedLoanType = Convert.ToString(oDT.GetValue("Name", 0)).Trim();
                            _MasterDBDataSource.SetValue("U_LoanType", 0, selectedLoanType);
                        }

                        if (pVal.ItemUID == "txtBankC")
                        {
                            string bankCode = Convert.ToString(oDT.GetValue("U_BankCode", 0)).Trim();
                            string bankName = Convert.ToString(oDT.GetValue("U_BankName", 0)).Trim();
                            _MasterDBDataSource.SetValue("U_BankCode", 0, bankCode);
                            _MasterDBDataSource.SetValue("U_BankName", 0, bankName);
                        }
                     break;

                    case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                        if (pVal.ItemUID == "txtBaseR" || pVal.ItemUID == "txtPremR")
                        {
                            CalculateTotalInterest();
                        }

                        if (pVal.ItemUID == "txtEffD" || pVal.ItemUID == "txtLoanT" || pVal.ItemUID == "txtBankC")
                        {
                            //FillMatrix();
                        }

                        if (pVal.ItemUID == "mtxInt" && (pVal.ColUID == "ColBaseR" || pVal.ColUID == "ColPremR"))
                        {
                            if (!pVal.BeforeAction && pVal.Row > 0)
                            {
                                CalculateRowInterest(pVal.Row);
                            }
                        }

                        if (pVal.ItemUID == "txtEffD" && !pVal.BeforeAction)
                        {
                            string effDateStr = _MasterDBDataSource.GetValue("U_EffectiveDate", 0).Trim();
                            if (!string.IsNullOrEmpty(effDateStr))
                            {
                                HandleEffectiveDateChange(effDateStr);
                            }
                        }

                        break;

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if(pVal.ItemUID == "btnPost")
                        {
                            PostInterestRateChanges();
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

                                FillMatrix();
                            }
                        }
                        break;

                }
            }
        }
        #endregion

        #region FormDefault
        public void FormDefault()
        {
            _Form.Freeze(true);
            try
            {
                _Form.DataBrowser.BrowseBy = "txtDocNo";
                _Form.Settings.Enabled = true;

                _MasterDBDataSource = _Form.DataSources.DBDataSources.Item("@INTERESTRATEH");
                _ChildDBDataSource = _Form.DataSources.DBDataSources.Item("@INTERESTRATEC");

                PopulateFiscalYears(_Form);
                //if (_Form.Mode == BoFormMode.fm_ADD_MODE)
                if (_Form.Mode == BoFormMode.fm_OK_MODE)
                {
                    string currentDate = DateTime.Today.ToString("yyyyMMdd");
                    _MasterDBDataSource.SetValue("U_PostingDate", 0, currentDate);

                    _MasterDBDataSource.SetValue("U_BaseInterest", 0, "0.0");
                    _MasterDBDataSource.SetValue("U_PremInterest", 0, "0.0");
                    _MasterDBDataSource.SetValue("U_RevInterest", 0, "0.0");


                    SAPbobsCOM.Recordset rsGen = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    try
                    {
                        string sql = @"SELECT MAX(CAST(IFNULL(""U_DocNo"", '0') AS INT)) AS ""MaxNo"" FROM ""@INTERESTRATEH""";

                        rsGen.DoQuery(sql);

                        int nextNo = 1;
                        if (!rsGen.EoF)
                        {
                            int currentMax = Convert.ToInt32(rsGen.Fields.Item("MaxNo").Value);
                            nextNo = currentMax + 1;
                        }

                        _MasterDBDataSource.SetValue("U_DocNo", 0, nextNo.ToString());
                    }
                    catch (Exception ex)
                    {
                        _MasterDBDataSource.SetValue("U_DocNo", 0, "1"); // Fallback if query fails
                    }
                    finally
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(rsGen);
                    }
                }
            }
            finally
            {
                _Form.Freeze(false);
            }
        }
        #endregion

        #region Calculate Revision Interest on the basis of (base + premium) interest
        private void CalculateTotalInterest()
        {
            try
            {
                _Form.Freeze(true);
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)_Form.Items.Item("mtxInt").Specific;

                double baseInt = 0;
                double premInt = 0;

                // Get values from Master DataSource (Header)
                double.TryParse(_MasterDBDataSource.GetValue("U_BaseInterest", 0).Trim(), out baseInt);
                double.TryParse(_MasterDBDataSource.GetValue("U_PremInterest", 0).Trim(), out premInt);

                double totalInt = baseInt + premInt;

                _MasterDBDataSource.SetValue("U_RevInterest", 0, totalInt.ToString());

                // Update all rows in the Child DataSource
                for (int i = 0; i < _ChildDBDataSource.Size; i++)
                {
                    // Only update rows that have a linked Loan (DocEntry)
                    string rowDocEntry = _ChildDBDataSource.GetValue("U_DocEntry", i).Trim();

                    if (!string.IsNullOrEmpty(rowDocEntry))
                    {
                        _ChildDBDataSource.SetValue("U_BaseRate", i, baseInt.ToString());
                        _ChildDBDataSource.SetValue("U_PremRate", i, premInt.ToString());
                        _ChildDBDataSource.SetValue("U_RevInterest", i, totalInt.ToString());
                    }
                }

                // Load the changes into the Matrix UI
                oMatrix.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.StatusBar.SetText("Calc Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
            }
            finally
            {
                _Form.Freeze(false);
            }
        }
        #endregion

        #region Fill the Matrix 
        private void FillMatrix()
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)_Form.Items.Item("mtxInt").Specific; 
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                _Form.Freeze(true);

                string lType = _MasterDBDataSource.GetValue("U_LoanType", 0).Trim();
                string bCode = _MasterDBDataSource.GetValue("U_BankCode", 0).Trim();
                string effDate = _MasterDBDataSource.GetValue("U_EffectiveDate", 0).Trim();

                if (string.IsNullOrEmpty(lType) || string.IsNullOrEmpty(bCode)) //|| string.IsNullOrEmpty(effDate)
                {
                    return;
                }

                string query = $@"SELECT * FROM (
                            SELECT
                                T0.""U_StartDate"", T0.""U_EndDate"", T3.""U_ToDate"", T3.""U_InterestPer"", 
                                T2.""U_LoanCode"", T0.""U_LoanNo"", T0.""U_BankFacNo"", T3.""DocEntry"", T3.""U_BaseInterestPer"", T3.""U_PremInterestPer"",
                                ROW_NUMBER() OVER (PARTITION BY T2.""U_LoanCode"" ORDER BY T3.""U_ToDate"" DESC) AS RN
                            FROM ""@LOANMASTERC"" T0
                            INNER JOIN ""@LOANMASTERH"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                            INNER JOIN ""@LOANMASTERDETAILSH"" T2 ON T0.""U_LoanCode"" = T2.""U_LoanCode""
                            INNER JOIN ""@LOANMASTERDETAILSC1"" T3 ON T2.""DocEntry"" = T3.""DocEntry""
                            WHERE T0.""U_LoanType"" = '{lType}' 
                            AND T1.""U_BankCode"" = '{bCode}' 
                            AND T3.""U_ToDate"" IS NOT NULL AND T0.""U_EndDate"" >= TO_DATE('{effDate}', 'yyyyMMdd')
                        ) X WHERE RN = 1 ORDER BY X.""U_LoanCode"" ASC ";

                rs.DoQuery(query);

                // Clear the Child DataSource and Matrix
                _ChildDBDataSource.Clear();
                oMatrix.Clear();

                int i = 0;
                while (!rs.EoF)
                {
                    _ChildDBDataSource.InsertRecord(i);

                    _ChildDBDataSource.SetValue("U_Check", i, "N");

                    _ChildDBDataSource.SetValue("U_DocEntry", i, rs.Fields.Item("DocEntry").Value.ToString());
                    _ChildDBDataSource.SetValue("U_LoanACCode", i, rs.Fields.Item("U_LoanCode").Value.ToString());
                    _ChildDBDataSource.SetValue("U_LoanACName", i, rs.Fields.Item("U_LoanNo").Value.ToString());
                    _ChildDBDataSource.SetValue("U_StartDate", i, ((DateTime)rs.Fields.Item("U_StartDate").Value).ToString("yyyyMMdd"));
                    _ChildDBDataSource.SetValue("U_EndDate", i, ((DateTime)rs.Fields.Item("U_EndDate").Value).ToString("yyyyMMdd"));
                    _ChildDBDataSource.SetValue("U_FacilityNo", i, rs.Fields.Item("U_BankFacNo").Value.ToString());
                    _ChildDBDataSource.SetValue("U_LastRevDate", i, ((DateTime)rs.Fields.Item("U_ToDate").Value).ToString("yyyyMMdd"));
                    _ChildDBDataSource.SetValue("U_PrevInterest", i, rs.Fields.Item("U_InterestPer").Value.ToString());

                    _ChildDBDataSource.SetValue("U_BaseRate", i, rs.Fields.Item("U_BaseInterestPer").Value.ToString());
                    _ChildDBDataSource.SetValue("U_PremRate", i, rs.Fields.Item("U_PremInterestPer").Value.ToString());

                    double baseInt = 0;
                    double premInt = 0;

                    // Parse UI values
                    double.TryParse(rs.Fields.Item("U_BaseInterestPer").Value, out baseInt);
                    double.TryParse(rs.Fields.Item("U_PremInterestPer").Value, out premInt);

                    double totalInt = baseInt + premInt;

                    _ChildDBDataSource.SetValue("U_RevInterest", i, totalInt.ToString());

                    // Default values for new rates from Header
                    //string baseRate = _MasterDBDataSource.GetValue("U_BaseInterest", 0).Trim();
                    //string premRate = _MasterDBDataSource.GetValue("U_PremInterest", 0).Trim();
                    //string revInt = _MasterDBDataSource.GetValue("U_RevInterest", 0).Trim();

                    //_ChildDBDataSource.SetValue("U_BaseRate", i, baseRate);
                    //_ChildDBDataSource.SetValue("U_PremRate", i, premRate);
                    //_ChildDBDataSource.SetValue("U_RevInterest", i, revInt);

                    rs.MoveNext();
                    i++;
                }

                // Load data into Matrix
                oMatrix.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.SetStatusBarMessage("Error filling matrix: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                _Form.Freeze(false);
            }
        }
        #endregion


        #region Calculate Row Interest
        private void CalculateRowInterest(int rowIdx)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)_Form.Items.Item("mtxInt").Specific;
            try
            {
                _Form.Freeze(true);

                // Get the values directly from the Matrix UI cells for immediate feedback
                SAPbouiCOM.EditText oEditBase = (SAPbouiCOM.EditText)oMatrix.Columns.Item("ColBaseR").Cells.Item(rowIdx).Specific;
                SAPbouiCOM.EditText oEditPrem = (SAPbouiCOM.EditText)oMatrix.Columns.Item("ColPremR").Cells.Item(rowIdx).Specific;
                SAPbouiCOM.EditText oEditRev = (SAPbouiCOM.EditText)oMatrix.Columns.Item("ColRInt").Cells.Item(rowIdx).Specific;

                double baseInt = 0;
                double premInt = 0;

                // Parse UI values
                double.TryParse(oEditBase.Value, out baseInt);
                double.TryParse(oEditPrem.Value, out premInt);

                double totalInt = baseInt + premInt;

                // Set the calculated value back to the UI EditText
                oEditRev.Value = totalInt.ToString();

                // Update the DataSource so the data is saved correctly (0-based index)
                _ChildDBDataSource.SetValue("U_RevInterest", rowIdx - 1, totalInt.ToString());
                _ChildDBDataSource.SetValue("U_BaseRate", rowIdx - 1, baseInt.ToString());
                _ChildDBDataSource.SetValue("U_PremRate", rowIdx - 1, premInt.ToString());
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.StatusBar.SetText("Row Calc Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
            }
            finally
            {
                _Form.Freeze(false);
            }
        }
        #endregion

        #region Posting Logic For Interest Rate Change

        private void PostInterestRateChanges()
        {
            SAPbobsCOM.GeneralService oGeneralService;
            SAPbobsCOM.GeneralData oGeneralData;
            SAPbobsCOM.GeneralDataParams oGeneralParams;
            SAPbobsCOM.CompanyService oCompService = Utilities.Application.Company.GetCompanyService();

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)_Form.Items.Item("mtxInt").Specific;

            oGeneralService = oCompService.GetGeneralService("LOANMASTERDETAILS");

            try
            {
                _Form.Freeze(true);

                oMatrix.FlushToDataSource();

                string hBaseStr = _MasterDBDataSource.GetValue("U_BaseInterest", 0).Trim();
                string hPremStr = _MasterDBDataSource.GetValue("U_PremInterest", 0).Trim();
                string hRevStr = _MasterDBDataSource.GetValue("U_RevInterest", 0).Trim();

                double hBase = 0, hPrem = 0, hTotal = 0;
                double.TryParse(hBaseStr, out hBase);
                double.TryParse(hPremStr, out hPrem);
                double.TryParse(hRevStr, out hTotal);

                // If all fields are zero or blank, stop the process
                if (hBase == 0 && hPrem == 0 && hTotal == 0)
                {
                    Utilities.Application.SBO_Application.MessageBox("Posting Failed: Header Interest Rates (Base, Premium, and Total) cannot be blank or zero.");
                    return;
                }


                //Get Global Dates from Header
                string hEffDateStr = _MasterDBDataSource.GetValue("U_EffectiveDate", 0).Trim();
                string hToDateStr = _MasterDBDataSource.GetValue("U_ToDate", 0).Trim();

                if (string.IsNullOrEmpty(hEffDateStr) || string.IsNullOrEmpty(hToDateStr))
                {
                    Utilities.Application.SBO_Application.MessageBox("Please select Effective Date and To Date in the header.");
                    return;
                }

                DateTime headerFromDate = DateTime.ParseExact(hEffDateStr, "yyyyMMdd", null);
                DateTime headerToDate = DateTime.ParseExact(hToDateStr, "yyyyMMdd", null);

                // PRE-VALIDATION: Check for Gaps in all selected rows before starting any updates
                bool anyRowSelected = false;
                for (int i = 0; i < _ChildDBDataSource.Size; i++)
                {
                    string isChecked = _ChildDBDataSource.GetValue("U_Check", i).Trim();
                    if (isChecked != "Y") continue;

                    anyRowSelected = true;

                    string docEntry = _ChildDBDataSource.GetValue("U_DocEntry", i).Trim();
                    string lastRevDateStr = _ChildDBDataSource.GetValue("U_LastRevDate", i).Trim();
                    string loanEndDateStr = _ChildDBDataSource.GetValue("U_EndDate", i).Trim();
                    string loanCode = _ChildDBDataSource.GetValue("U_LoanACCode", i).Trim();


                    if (string.IsNullOrEmpty(docEntry) || string.IsNullOrEmpty(lastRevDateStr)) continue;

                    DateTime lastRevDate = DateTime.ParseExact(lastRevDateStr, "yyyyMMdd", null);
                    DateTime loanEndDate = DateTime.ParseExact(loanEndDateStr, "yyyyMMdd", null);
                    DateTime expectedFromDate = lastRevDate.AddDays(1);

                    // Check if Header Effective Date matches Last Rev Date + 1
                    if (headerFromDate != expectedFromDate)
                    {
                        Utilities.Application.SBO_Application.MessageBox(
                            $"Gap Detected for Loan Code: {loanCode} (DocEntry: {docEntry})!\n" +
                            $"Expected Effective Date: {expectedFromDate:yyyy-MM-dd}\n" +
                            $"Selected Effective Date: {headerFromDate:yyyy-MM-dd}");
                        return; // Stop the whole process if a gap is found
                    }

                    if (headerFromDate > loanEndDate)
                    {
                        Utilities.Application.SBO_Application.MessageBox($"Loan {loanCode} is already expired on {loanEndDate:yyyy-MM-dd}. Cannot post new rates.");
                        return;
                    }
                }

                if (!anyRowSelected)
                {
                    Utilities.Application.SBO_Application.MessageBox("No rows selected for posting.");
                    return;
                }

                // POSTING: If all rows pass validation, proceed to update
                int successCount = 0;
                for (int i = 0; i < _ChildDBDataSource.Size; i++)
                {
                    string isChecked = _ChildDBDataSource.GetValue("U_Check", i).Trim();
                    if (isChecked != "Y") continue;

                    string docEntry = _ChildDBDataSource.GetValue("U_DocEntry", i).Trim();
                    string loanEndDateStr = _ChildDBDataSource.GetValue("U_EndDate", i).Trim();
                    string rowBase = _ChildDBDataSource.GetValue("U_BaseRate", i).Trim();
                    string rowPrem = _ChildDBDataSource.GetValue("U_PremRate", i).Trim();
                    string rowRev = _ChildDBDataSource.GetValue("U_RevInterest", i).Trim();

                    DateTime loanEndDate = DateTime.ParseExact(loanEndDateStr, "yyyyMMdd", null);

                    DateTime actualToDate = (headerToDate > loanEndDate) ? loanEndDate : headerToDate;

                    if (string.IsNullOrEmpty(docEntry)) continue;

                    oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                    oGeneralParams.SetProperty("DocEntry", docEntry);
                    oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                    SAPbobsCOM.GeneralDataCollection oSlabs = oGeneralData.Child("LOANMASTERDETAILSC1");
                    SAPbobsCOM.GeneralData oRowToUpdate = null;

                    // Identify the last row for blank handling
                    if (oSlabs.Count > 0)
                    {
                        SAPbobsCOM.GeneralData oLastRow = oSlabs.Item(oSlabs.Count - 1);
                        if (IsDateEmpty(oLastRow.GetProperty("U_FromDate")))
                        {
                            oRowToUpdate = oLastRow;
                        }
                    }

                    if (oRowToUpdate == null)
                    {
                        oRowToUpdate = oSlabs.Add();
                    }

                    // Set the Dates from Header
                    oRowToUpdate.SetProperty("U_FromDate", headerFromDate);
                    oRowToUpdate.SetProperty("U_ToDate", actualToDate);

                    // Set Rates from Matrix Row
                    oRowToUpdate.SetProperty("U_BaseInterestPer", rowBase);
                    oRowToUpdate.SetProperty("U_PremInterestPer", rowPrem);
                    oRowToUpdate.SetProperty("U_InterestPer", rowRev);

                    // No. of Days calculation based on Header Dates
                    int noOfDays = (actualToDate - headerFromDate).Days + 1;
                    oRowToUpdate.SetProperty("U_NoDays", noOfDays);

                    // Nepali Dates based on Header Dates
                    oRowToUpdate.SetProperty("U_NFromDate", GetNepaliDate(headerFromDate));
                    oRowToUpdate.SetProperty("U_NToDate", GetNepaliDate(actualToDate));

                    // Add the blank row for LoanMasterDetails UI readiness
                    oSlabs.Add();

                    // Update Loan Master Details record
                    oGeneralService.Update(oGeneralData);
                    successCount++;
                }

                if (successCount > 0)
                {
                    try
                    {
                        Utilities.Application.SBO_Application.MessageBox($"Successfully updated {successCount} loans.");
                        _Form.Mode = BoFormMode.fm_OK_MODE;
                        _Form.Close();
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.SetStatusBarMessage("Posting Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
            }
            finally
            {
                try
                {
                    _Form.Freeze(false);
                }
                catch
                {

                }
            }
        }

        #endregion

        #region Check if date is empty
        private bool IsDateEmpty(object dateObj)
        {
            if (dateObj == null) return true;

            if (dateObj is DateTime dt)
            {
                if (dt.Year < 1920 || dt == DateTime.MinValue)
                    return true;
            }

            string dateStr = dateObj.ToString();
            if (string.IsNullOrWhiteSpace(dateStr) || dateStr.Contains("1899") || dateStr.Contains("1900"))
                return true;

            return false;
        }
        #endregion


        #region Get Nepali Date
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


        #region Date Calculation
        private void HandleEffectiveDateChange(string effectiveDateStr)
        {
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                _Form.Freeze(true);
                DateTime effDate = DateTime.ParseExact(effectiveDateStr, "yyyyMMdd", null);

                // Get Nepali Date for the Effective Date
                string nepEffDate = GetNepaliDate(effDate);
                _MasterDBDataSource.SetValue("U_NEffectiveDate", 0, nepEffDate);

                // Find the end of this Nepali Month
                // Extract Year and Month (e.g., from "2080/05/12" to "2080/05")
                if (nepEffDate.Length >= 7)
                {
                    string yrMonthPrefix = nepEffDate.Substring(0, 7);

                    // Query to find the maximum (last) date in that Nepali Month
                    string sql = $@"SELECT MAX(""U_EnglishDate"") as ""MaxEng"", MAX(""U_NepaliDate"") as ""MaxNep""
                            FROM ""@NEPALICALENDER""
                            WHERE ""U_NepaliDate"" LIKE '{yrMonthPrefix}%'";

                    rs.DoQuery(sql);

                    if (!rs.EoF)
                    {
                        DateTime engToDate = (DateTime)rs.Fields.Item("MaxEng").Value;
                        string nepToDate = rs.Fields.Item("MaxNep").Value.ToString();

                        // Set English To Date (Item_10 / U_ToDate)
                        _MasterDBDataSource.SetValue("U_ToDate", 0, engToDate.ToString("yyyyMMdd"));

                        // Set Nepali To Date (txtNTo / U_NToDate)
                        _MasterDBDataSource.SetValue("U_NToDate", 0, nepToDate);
                    }
                }
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.StatusBar.SetText("Date Calculation Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                _Form.Freeze(false);
            }
        }
        #endregion

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
                    _MasterDBDataSource.SetValue("U_EffectiveDate", 0, sDate.ToString("yyyyMMdd"));
                    _MasterDBDataSource.SetValue("U_ToDate", 0, eDate.ToString("yyyyMMdd"));

                    // Force the UI to refresh from the DataSource
                    ((SAPbouiCOM.EditText)oForm.Items.Item("txtEffD").Specific).Value = sDate.ToString("yyyyMMdd");
                    ((SAPbouiCOM.EditText)oForm.Items.Item("Item_10").Specific).Value = eDate.ToString("yyyyMMdd");

                    HandleEffectiveDateChange(sDate.ToString("yyyyMMdd"));
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

    }
}
