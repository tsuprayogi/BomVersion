using SAPbouiCOM;

namespace BOM_Version.Services
{
    public class ModeHandlerService
    {
        private readonly IForm _form;
        private readonly MatrixService _matrix;
        private readonly EditText _editItemCode;
        private readonly StaticText _lblItemName;
        private readonly ComboBox _combo;
        private readonly Button _btnSave;

        public bool IsFindMode { get; private set; }

        public ModeHandlerService(
            IForm form,
            MatrixService matrix,
            EditText ic,
            StaticText nm,
            ComboBox ver,
            Button save)
        {
            _form = form;
            _matrix = matrix;
            _editItemCode = ic;
            _lblItemName = nm;
            _combo = ver;
            _btnSave = save;
        }

        // ================================================
        // ADD MODE
        // ================================================
        public void SetAddMode()
        {
            _form.Mode = BoFormMode.fm_ADD_MODE;
            IsFindMode = false;

            _btnSave.Caption = "Add";

            ClearHeader();
            ClearDetail();
            _matrix.AddEmptyRow();
        }

        // ================================================
        // FIND MODE
        // ================================================
        public void SetFindMode()
        {
            _form.Mode = BoFormMode.fm_FIND_MODE;
            IsFindMode = true;

            _btnSave.Caption = "Find";

            ClearHeader();
            ClearDetail();
        }

        // ================================================
        // OK MODE
        // ================================================
        public void SetOKMode()
        {
            _form.Mode = BoFormMode.fm_OK_MODE;
            IsFindMode = false;

            _btnSave.Caption = "OK";
        }

        // ================================================
        // UPDATE MODE (ketika field berubah)
        // ================================================
        public void SwitchToUpdateModeIfOK()
        {
            if (_btnSave.Caption == "OK")
                _btnSave.Caption = "Update";
        }

        // ================================================
        // Reset data
        // ================================================
        private void ClearHeader()
        {
            _editItemCode.Value = "";
            _lblItemName.Caption = "";
            try { _combo.Select(0, BoSearchKey.psk_Index); } catch { }
        }

        private void ClearDetail()
        {
            var ds = _form.DataSources.DBDataSources.Item("@PRODBOM_L");
            ds.Clear();
            _matrix.Reload();
        }
    }
}
