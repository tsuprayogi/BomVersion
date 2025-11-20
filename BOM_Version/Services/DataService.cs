using System;
using SAPbouiCOM;
using SAPbobsCOM;
using BOM_Version.Repositories;

namespace BOM_Version.Services
{
    public class DataService : IDataService
    {
        private readonly IForm _form;
        private readonly Matrix _matrix;
        private readonly EditText _editItemCode;
        private readonly ComboBox _comboVersion;
        private readonly DocumentRepository _repo;
        private readonly ValidationService _validator;

        private const string FIELD_COMP = "U_U_Component";
        private const string FIELD_NAME = "U_U_ItemName";
        private const string FIELD_QTY = "U_U_Quantity";

        public DataService(IForm form, Matrix matrix, EditText item, ComboBox version)
        {
            _form = form;
            _matrix = matrix;
            _editItemCode = item;
            _comboVersion = version;

            _repo = new DocumentRepository();
            _validator = new ValidationService();
        }


        // Helper di kelas DataService (atau util umum)
        private void WithFreeze(Action action)
        {
            var realForm = _form as Form;
            if (realForm == null)
            {
                action();
                return;
            }

            try
            {
                realForm.Freeze(true);
                action();
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // optional: log
                throw;
            }
            finally
            {
                try { realForm.Freeze(false); } catch { /* ignore */ }
            }
        }

        private string Esc(string t) => _repo.Escape(t);

        // ===========================================================
        // FIND LAST RECORD BY ITEMCODE
        // ===========================================================
        public string FindLastByItemCode(string itemCode)
        {
            string q = $@"
                SELECT TOP 1 Code 
                FROM [@PRODBOM]
                WHERE U_U_ItemCode = '{Esc(itemCode)}'
                ORDER BY Code DESC";

            var rs = _repo.Exec(q);
            if (rs.RecordCount == 0) return null;
            return rs.Fields.Item("Code").Value.ToString();
        }

        // ===========================================================
        // LOAD HEADER
        // ===========================================================
        public bool LoadHeader(string code)
        {
            string q = $"SELECT * FROM [@PRODBOM] WHERE Code = '{Esc(code)}'";
            var rs = _repo.Exec(q);

            if (rs.RecordCount == 0) return false;

            _editItemCode.Value = rs.Fields.Item("U_U_ItemCode").Value.ToString();
            string ver = rs.Fields.Item("U_U_Version").Value?.ToString();

            try
            {
                _comboVersion.Select(ver, BoSearchKey.psk_ByValue);
            }
            catch { }

            return true;
        }



        public bool LoadDetail(string codeHeader)
        {
            string q = $@"
                SELECT U_U_Component, U_U_ItemName, U_U_Quantity
                FROM [@PRODBOM_L]
                WHERE Code LIKE '{Esc(codeHeader)}-%'
                ORDER BY Code";
            var rs = _repo.Exec(q);
            var ds = _form.DataSources.DBDataSources.Item("@PRODBOM_L");

            WithFreeze(() =>
            {
                ds.Clear();
                int row = 0;
                while (!rs.EoF)
                {
                    ds.InsertRecord(row);
                    ds.SetValue(FIELD_COMP, row, rs.Fields.Item("U_U_Component").Value.ToString());
                    ds.SetValue(FIELD_NAME, row, rs.Fields.Item("U_U_ItemName").Value.ToString());
                    ds.SetValue(FIELD_QTY, row, rs.Fields.Item("U_U_Quantity").Value.ToString());
                    row++;
                    rs.MoveNext();
                }
                _matrix.LoadFromDataSource();
            });

            return true;
        }


        public void InitComboBoxVersion()
        {
            if (_comboVersion == null) return;

            try
            {
                // Hapus semua valid values (loop mundur lebih aman)
                for (int i = _comboVersion.ValidValues.Count - 1; i >= 0; i--)
                    _comboVersion.ValidValues.Remove(i, BoSearchKey.psk_Index);

                // Tambah nilai versi 1..5
                for (int i = 1; i <= 5; i++)
                    _comboVersion.ValidValues.Add(i.ToString(), i.ToString());

                // Pilih default (index 0)
                if (_comboVersion.ValidValues.Count > 0)
                    _comboVersion.Select(0, BoSearchKey.psk_Index);
            }
            catch
            {
                // optional: log error
            }
        }

        // ===========================================================
        // SAVE
        // ===========================================================
        public bool Save(out string message)
        {
            try
            {
                _matrix.FlushToDataSource();
                var ds = _form.DataSources.DBDataSources.Item("@PRODBOM_L");

                if (!_validator.ValidateHeader(_editItemCode))
                {
                    message = "Header tidak valid.";
                    return false;
                }
                if (!_validator.ValidateDetail(ds, FIELD_COMP))
                {
                    message = "Detail tidak valid.";
                    return false;
                }

                string codeHeader = $"BOM-{DateTime.Now:yyyyMMddHHmmss}";
                string qh = $@"INSERT INTO [@PRODBOM] (Code,Name,U_U_ItemCode,U_U_Version)
                       VALUES ('{codeHeader}','{codeHeader}','{Esc(_editItemCode.Value.Trim())}','{Esc(_comboVersion.Value.Trim())}')";

                WithFreeze(() => { _repo.Execute(qh); });

                int idx = 1;
                for (int i = 0; i < ds.Size; i++)
                {
                    string comp = ds.GetValue(FIELD_COMP, i).Trim();
                    if (string.IsNullOrEmpty(comp)) continue;

                    string qd = $@"INSERT INTO [@PRODBOM_L] (Code,Name,{FIELD_COMP},{FIELD_NAME},{FIELD_QTY})
                           VALUES ('{codeHeader}-{idx}','{codeHeader}-{idx}','{Esc(ds.GetValue(FIELD_COMP, i).Trim())}',
                                   '{Esc(ds.GetValue(FIELD_NAME, i).Trim())}','{Esc(ds.GetValue(FIELD_QTY, i).Trim())}')";
                    _repo.Execute(qd);
                    idx++;
                }

                message = "Data berhasil disimpan!";
                return true;
            }
            catch (Exception ex)
            {
                message = "Error saat menyimpan: " + ex.Message;
                return false;
            }
        }

        // ===========================================================
        // UPDATE
        // ===========================================================
        public void Update(string codeHeader)
        {
            _matrix.FlushToDataSource();

            if (!_validator.ValidateHeader(_editItemCode)) return;

            var company = (SAPbobsCOM.Company)Program.SBO_Application.Company.GetDICompany();
            var oHeader = company.UserTables.Item("PRODBOM");

            if (!oHeader.GetByKey(codeHeader))
            {
                Program.SBO_Application.MessageBox("Header tidak ditemukan.");
                return;
            }

            // update header
            oHeader.UserFields.Fields.Item("U_U_ItemCode").Value = _editItemCode.Value.Trim();
            oHeader.UserFields.Fields.Item("U_U_Version").Value = _comboVersion.Value.Trim();

            if (oHeader.Update() != 0)
            {
                Program.SBO_Application.MessageBox("Header update error: " + company.GetLastErrorDescription());
                return;
            }

            // delete old detail
            _repo.Execute($"DELETE FROM [@PRODBOM_L] WHERE Code LIKE '{Esc(codeHeader)}-%'");

            // insert new detail
            int idx = 1;
            for (int row = 1; row <= _matrix.RowCount; row++)
            {
                string comp = (_matrix.GetCellSpecific("DetCode", row) as EditText)?.Value.Trim();
                if (string.IsNullOrEmpty(comp)) continue;

                string name = Esc((_matrix.GetCellSpecific("DetName", row) as EditText)?.Value.Trim());
                string qty = Esc((_matrix.GetCellSpecific("colQty", row) as EditText)?.Value.Trim() ?? "0");

                string qd = $@"
                    INSERT INTO [@PRODBOM_L]
                    (Code,Name,{FIELD_COMP},{FIELD_NAME},{FIELD_QTY})
                    VALUES
                    ('{codeHeader}-{idx}',
                     '{codeHeader}-{idx}',
                     '{Esc(comp)}',
                     '{name}',
                     '{qty}')
                ";
                _repo.Execute(qd);
                idx++;
            }

            Program.SBO_Application.MessageBox("Data berhasil di-update!");
        }
    }
}
