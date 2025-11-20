using SAPbouiCOM;

namespace BOM_Version.Services
{
    public class MatrixService
    {
        private readonly IForm _form;
        private readonly Matrix _matrix;

        private readonly string FIELD_COMP = "U_U_Component";
        private readonly string FIELD_NAME = "U_U_ItemName";
        private readonly string FIELD_QTY = "U_U_Quantity";

        public MatrixService(IForm form, Matrix matrix)
        {
            _form = form;
            _matrix = matrix;
        }

        private DBDataSource DS => _form.DataSources.DBDataSources.Item("@PRODBOM_L");

        private void WithFreeze(System.Action action)
        {
            Form realForm = _form as Form;
            if (realForm == null)
            {
                // jika bukan Form, jalankan langsung
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
                // opsional: log atau ignore
                throw;
            }
            finally
            {
                try { realForm.Freeze(false); } catch { /* ignore */ }
            }
        }



        // ==========================================================
        // FLUSH & LOAD
        // ==========================================================
        public void Flush()
        {
            try { _matrix.FlushToDataSource(); } catch { }
        }

        public void Reload()
        {
            WithFreeze(() =>
            {
                _matrix.LoadFromDataSource();
                _matrix.AutoResizeColumns();
            });
        }


        // ==========================================================
        // ADD EMPTY ROW
        // ==========================================================
        public void AddEmptyRow()
        {
            Flush();

            int idx = DS.Size;
            DS.InsertRecord(idx);

            DS.SetValue(FIELD_COMP, idx, "");
            DS.SetValue(FIELD_NAME, idx, "");
            DS.SetValue(FIELD_QTY, idx, "");

            Reload();
        }

        // ==========================================================
        // DELETE ROW
        // ==========================================================
        public void DeleteRow(int row)
        {
            Flush();
            int index = row - 1;

            if (index < 0 || index >= DS.Size)
                return;

            DS.RemoveRecord(index);
            Reload();
        }

        // ==========================================================
        // SET CELL FROM CFL
        // ==========================================================
        public void ApplyCFLToRow(int row, string comp, string name)
        {
            if (row <= 0) return;

            Flush();

            int dsIndex = row - 1;

            DS.SetValue(FIELD_COMP, dsIndex, comp);
            DS.SetValue(FIELD_NAME, dsIndex, name);

            Reload();

            // jika baris terakhir terisi, tambahkan baris baru
            if (row == _matrix.RowCount)
                AddEmptyRow();
        }


        //GET VALUE CELL//
        public string GetCell(int row, string col)
        {
            var cell = _matrix.GetCellSpecific(col, row) as EditText;
            return cell?.Value?.Trim() ?? "";
        }

        // ==========================================================
        // VALIDATION CHECK
        // ==========================================================
        public bool HasAnyDetail()
        {
            Flush();

            for (int i = 0; i < DS.Size; i++)
            {
                string comp = DS.GetValue(FIELD_COMP, i).Trim();
                if (!string.IsNullOrEmpty(comp))
                    return true;
            }

            return false;
        }
    }
}
