using SAPbouiCOM;

namespace BOM_Version.Services
{
    public class CFLResult
    {
        public bool IsHandled { get; set; }
        public string ItemCode { get; set; }
        public string ItemName { get; set; }
    }

    public class CFLService
    {
        private readonly IForm _form;
        private readonly MatrixService _matrix;

        public CFLService(IForm form, MatrixService matrix)
        {
            _form = form;
            _matrix = matrix;
        }

        public CFLResult HandleMatrixCFL(int row, string col, string itemCode, string itemName)
        {
            if (col != "DetCode" || row <= 0)
                return new CFLResult { IsHandled = false };

            // Apply ke matrix + tambahkan row jika diperlukan
            _matrix.ApplyCFLToRow(row, itemCode, itemName);

            return new CFLResult
            {
                IsHandled = true,
                ItemCode = itemCode,
                ItemName = itemName
            };
        }

        public CFLResult HandleHeaderCFL(string itemCode, string itemName)
        {
            return new CFLResult
            {
                IsHandled = true,
                ItemCode = itemCode,
                ItemName = itemName
            };
        }

        public (string Code, string Name) ExtractSelected(ItemEvent pVal)
        {
            var cfl = (ChooseFromListEvent)pVal;
            var dt = cfl.SelectedObjects;

            if (dt == null || dt.Rows.Count == 0)
                return ("", "");

            return (
                dt.GetValue("ItemCode", 0).ToString(),
                dt.GetValue("ItemName", 0).ToString()
            );
        }
    }
}
