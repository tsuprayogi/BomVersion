using SAPbouiCOM.Framework;

namespace BOM_Version.Services
{
    public class ValidationService
    {
        public bool ValidateHeader(SAPbouiCOM.EditText itemCode)
        {
            if (string.IsNullOrWhiteSpace(itemCode.Value?.Trim()))
            {
               Application.SBO_Application.MessageBox("ItemCode tidak boleh kosong!");
                return false;
            }
            return true;
        }

        public bool ValidateDetail(SAPbouiCOM.DBDataSource ds, string fieldComp)
        {
            for (int i = 0; i < ds.Size; i++)
            {
                string comp = ds.GetValue(fieldComp, i).Trim();
                if (!string.IsNullOrEmpty(comp))
                    return true;
            }

            Application.SBO_Application.MessageBox("Detail minimal 1 row!");
            return false;
        }
    }
}
