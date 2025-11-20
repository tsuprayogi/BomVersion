using SAPbobsCOM;
using SAPbouiCOM.Framework;

namespace BOM_Version.Repositories
{
    public class DocumentRepository
    {
        private readonly Company _company;

        public DocumentRepository()
        {
            _company = (Company)Application.SBO_Application.Company.GetDICompany();
        }

        public Recordset Exec(string sql)
        {
            var rs = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
            rs.DoQuery(sql);
            return rs;
        }

        public void Execute(string sql)
        {
            var rs = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
            rs.DoQuery(sql);
        }

        public string Escape(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";
            return text.Replace("'", "''");
        }
    }
}
