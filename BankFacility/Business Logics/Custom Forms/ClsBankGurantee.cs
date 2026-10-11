using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Sap.Data.Hana;
using SAPbobsCOM;
using SAPbouiCOM;
namespace BankFacility
{
    class ClsBankGurantee :Base
    {
        #region VARIABLE DECLARATION

        SAPbouiCOM.Item _ExistItem, _NewItem, _Items;
        private bool flag;
        public SAPbouiCOM.Matrix _Matrix;
        SAPbouiCOM.StaticText _StaticText;
        SAPbouiCOM.Button _Button;
        string _DocEntry = "";
        CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep;
        string ReportPath = string.Empty;
        private static HanaConnection _conn;
        private static HanaDataAdapter _da;
        //private static DataTable _dataTable;
        private static HanaParameter _parm1;
        private static HanaParameter _parm2;
        private static int _errors = 0;
        string strcon, _Sql;
        private SAPbouiCOM.EditText _EditText;
        private SAPbouiCOM.ComboBox _ComboBoxM;
        private SAPbouiCOM.ComboBox _ComboBoxC;
        private SAPbouiCOM.ComboBox _ComboBoxB;
        private SAPbouiCOM.DBDataSource _MasterDBDataSource = null, _ChildDBDataSource = null;
        private int RowID { get; set; }
        private string _Code { get; set; }
        private string _LineId { get; set; }
        private string _ActivityType = string.Empty;
        string _ItemCode = string.Empty, _ItemName = string.Empty;
        private SAPbobsCOM.Recordset _Rs, _RS,_rs,_RSD;
        SAPbouiCOM.ChooseFromListEvent _SysCFLEvent = null;

        SAPbouiCOM.DBDataSource _DBData = null;
        private System.Collections.Generic.List<string> ListOfBlockedID;
        internal string _ServiceID { get; set; }
        internal string _ActivityTypes { get; set; }

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR
        public ClsBankGurantee()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsBankGurantee()
        {
        }
        #endregion

        #region FormDataEvent
        public override void FormData_Event(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool oBubbleEvent)
        {
            try
            {
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                        Form.Freeze(true);
                        Utilities.SerializedMartix(ref _Form, ref _Matrix);
                        Form.Freeze(false);
                        break;
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
        }
        #endregion

        #region ItemEvent
        public override void Item_Event(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            #region Before Action
            if (pVal.BeforeAction)
            {
                switch (pVal.EventType)
                {

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1" && Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {

                            _ChildDBDataSource.InsertRecord(_ChildDBDataSource.Size);   //Add new row on form nevigate
                            _Matrix.LoadFromDataSource();
                            Utilities.SerializedMartix(ref _Form, ref _Matrix);

                        }
                        break;

                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        if (pVal.ItemUID == "5" && pVal.CharPressed != 9 && pVal.InnerEvent == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        if (pVal.ItemUID == "6" && pVal.CharPressed != 9 && pVal.InnerEvent == false)
                        {
                            BubbleEvent = false;
                            return;
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
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        //if (pVal.ItemUID == "7" && pVal.ColUID == "V_7")
                        //{
                        //    //_ComboBoxC = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_7").Cells.Item(pVal.Row).Specific;
                        //    _ComboBoxC.Select(2, SAPbouiCOM.BoSearchKey.psk_Index);

                        //}
                        if(pVal.ItemUID=="7" && pVal.ColUID== "V_2")
                        {
                            _ComboBoxB = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific;
                            _EditText = (SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(pVal.Row).Specific;
                            _EditText.Value = _ComboBoxB.Selected.Description.ToString();
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1" && Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {

                            //_ChildDBDataSource.InsertRecord(_ChildDBDataSource.Size);   //Add new row on form nevigate
                            //_Matrix.LoadFromDataSource();
                            Utilities.SerializedMartix(ref _Form, ref _Matrix);
                            

                        }
                        break;

                }
            }

            #endregion

        }
        #endregion

        # region MENU EVENT
        public override void Menu_Event(ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                #region Before Action
                if (pVal.BeforeAction)
                {
                    switch (pVal.MenuUID)
                    {

                    }
                }
                #endregion

                #region After Action
                else
                {
                    switch (pVal.MenuUID)
                    {
                        //case Constants.System_Menus.mnu_DELETE_ROW:
                        //    if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_3").Cells.Item(_Matrix.VisualRowCount).Specific).Value.ToString() == "")
                        //    {

                        //        _Matrix.DeleteRow(_Matrix.VisualRowCount);
                        //        _Matrix.FlushToDataSource();

                        //    }
                        //    break;
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally { }
        }
        # endregion

        #region  Conditional Reload data in Form

        private void ConditionalReloadData()
        {
            try
            {
                _Form.Freeze(true);
                _DBData.Query(null);
                _Matrix.LoadFromDataSource();
                _DBData.Clear();
                _Matrix.AddRow(1, _Matrix.RowCount);

            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally { _Form.Freeze(false); }
        }

        #endregion

        #region FORM DEFAULT
        internal void FormDefault(string _BPCode, string _BPName)
        {
            try
            {
                _Form.Freeze(true);
                _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("7").Specific;
                _MasterDBDataSource = _Form.DataSources.DBDataSources.Item("@BGSETUPM");
                _ChildDBDataSource = _Form.DataSources.DBDataSources.Item("@BGSETUPC");
                // _MasterDBDataSource.SetValue("Code", 0, Utilities.getMaxColumnValueNum("@ITEMDUTYSETUPM", "Code"));
                _MasterDBDataSource.SetValue("Code", 0, _BPCode);
                _Matrix.AddRow(1, -1);
                _MasterDBDataSource.SetValue("U_BPCode", 0, _BPCode);
                _MasterDBDataSource.SetValue("U_BPName", 0, _BPName);
                _ComboBoxM = ((SAPbouiCOM.ComboBox)_Form.Items.Item("1000003").Specific);
                _ComboBoxM.ValidValues.Add("Active","Active");
                _ComboBoxM.ValidValues.Add("InActive", "In Active");
                _ComboBoxC = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_7").Cells.Item(_Matrix.RowCount).Specific;
                _ComboBoxC.ValidValues.Add("Open","Open");
                _ComboBoxC.ValidValues.Add("CLose", "CLose");
                _ComboBoxC.ValidValues.Add("ReNew", "ReNew");
                _ComboBoxC.ValidValues.Add("Claim Processing", "Claim Processing");
                _ComboBoxC.ValidValues.Add("ReNew Processing", "ReNew Processing");
                _ComboBoxB = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_2").Cells.Item(_Matrix.RowCount).Specific;
                Utilities.FillCombo(ref _ComboBoxB, "select distinct \"BankCode\" ,\"BankName\" from ODSC");
                //((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific) = 
                Utilities.ExecuteSQL(ref _rs, "select T0.\"GroupName\",T1.\"Currency\" ," +
                "case when T1.\"validFor\" ='Y' then 'Active' else 'In Active' end as \"BP Status\" "+
                "from \"OCRG\" T0 inner join \"OCRD\" T1 on T1.\"GroupCode\" = T0.\"GroupCode\" "+
                "where T1.\"CardCode\"='" + _BPCode + "'");
                if(_rs.RecordCount>0)
                {
                    //string name = _rs.Fields.Item("GroupName").Value;
                    ((SAPbouiCOM.EditText)_Form.Items.Item("15").Specific).Value = _rs.Fields.Item("GroupName").Value;
                    ((SAPbouiCOM.EditText)_Form.Items.Item("18").Specific).Value = _rs.Fields.Item("Currency").Value;
                    
                    if(_rs.Fields.Item("BP Status").Value=="Active")
                    {
                        _ComboBoxM.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                    else
                    {
                        _ComboBoxM.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
               
                }
                Utilities.ExecuteSQL(ref _RSD, "select T1.\"U_BGEDate\" from \"@BGSETUPC\" T1 where T1.\"Code\"= '" + _BPCode + "'");
                
                    //for (int i = 0; i < _RSD.RecordCount; i++)
                    //{
                    //    if (_RSD.Fields.Item("U_BGEDate").Value < DateTime.Now)
                    //    {
                    //        _ComboBoxC.Select("Terminated", BoSearchKey.psk_ByValue);
                    //    }
                    //}
                
                Utilities.ExecuteSQL(ref _RS, "Select \"Code\" From \"@BGSETUPM\" Where \"U_BPCode\" = '" + _BPCode + "'");
                if (_RS.RecordCount > 0)
                {
                    _Form.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    ((SAPbouiCOM.EditText)_Form.Items.Item("5").Specific).Value = _BPCode.ToString().Trim();
                    //((SAPbouiCOM.EditText)_Form.Items.Item("6").Specific).Value = ItemName.ToString();
                    _Form.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    if (_ChildDBDataSource.GetValue("U_BGStatus", _ChildDBDataSource.Size - 1).Trim() != "")
                    {
                        Form.Freeze(true);
                        _ChildDBDataSource.InsertRecord(_ChildDBDataSource.Size);   //Add new row on form nevigate
                        _Matrix.LoadFromDataSource();
                        // Utilities.SerializedMartix(ref _Form, ref _Matrix);
                        Form.Freeze(false);
                    }
                }
                Utilities.SerializedMartix(ref _Form, ref _Matrix);


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

        #region Set New line
        private void SetNewLine(string FormUID, int Row)
        {
            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("7").Specific;
            _ChildDBDataSource.Offset = Row;
            _ChildDBDataSource.SetValue("LineId", _ChildDBDataSource.Offset, Row.ToString());
            _ChildDBDataSource.SetValue("U_BGNo", _ChildDBDataSource.Offset, "");
            _ChildDBDataSource.SetValue("U_BankCod", _ChildDBDataSource.Offset, "");
            _ChildDBDataSource.SetValue("U_BankNam", _ChildDBDataSource.Offset, "");
            _ChildDBDataSource.SetValue("U_BGSDate", _ChildDBDataSource.Offset, "");
            _ChildDBDataSource.SetValue("U_BGEDate", _ChildDBDataSource.Offset, "");
            _ChildDBDataSource.SetValue("U_BGLimit", _ChildDBDataSource.Offset, "");
            _ChildDBDataSource.SetValue("U_BGStatus", _ChildDBDataSource.Offset, "");
            _Matrix.SetLineData(Row);
            _Matrix.FlushToDataSource();
            Utilities.SerializedMartix(ref _Form, ref _Matrix);
        }
        #endregion
    }
}
