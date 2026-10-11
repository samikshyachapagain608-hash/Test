using BankFacility;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BankFacilit
{
    class ClsLoanDimension : Base
    {
        public ClsLoanDimension() { }
        ~ClsLoanDimension() { }


        public override void RightClick_Event(ref ContextMenuInfo oEventInfo, ref bool oBubbleEvent)
        {
            Utilities.Application.SBO_Application.Menus.Item("1283").Enabled = false;

        }

        internal void FormDefault()
        {
            Utilities.Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "LOANDIM", "");
            SAPbouiCOM.Form oForm = Utilities.Application.SBO_Application.Forms.ActiveForm;
            DisableColumns(oForm);
        }

        internal void DisableColumns(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("3").Specific;

                oMatrix.Columns.Item("U_DimCode").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                oMatrix.Columns.Item("Code").Editable = false;
                oMatrix.Columns.Item("Name").Editable = false;
                oMatrix.Columns.Item("U_ShortName").Editable = false;
                oMatrix.Columns.Item("U_NextNo").Editable = false;
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.SetStatusBarMessage(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }
    }
}
