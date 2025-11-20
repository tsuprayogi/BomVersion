using SAPbouiCOM;
using SAPbouiCOM.Framework;
using BOM_Version.Services;
using BOM_Version.Repositories;



namespace BOM_Version
{
    [FormAttribute("BOM_Version.Alternate", "Alternate.b1f")]
    public class Alternate : UserFormBase
    {
        // ===== UI =====
        private Matrix matrix;
        private EditText edtItem;
        private StaticText lblName;
        private ComboBox cboVersion;
        private Button btnSave;
        private Button btnClose;

        // ===== Services =====
        private MatrixService matrixService;
        private ModeHandlerService modeService;
        private CFLService cflService;
        private MenuService menuService;
        private DataService dataService;
        private NavigationService navService;

        // ===== CFL state =====
        private int lastRow = -1;
        private string lastCol = "";

        public Alternate()
        {
           Program.SBO_Application.ItemEvent += OnItemEvent;
            Program.SBO_Application.MenuEvent += OnMenuEvent;
        }

        public override void OnInitializeComponent()
        {
            // Ambil UI
            matrix = GetItem("Item_1").Specific as Matrix;
            edtItem = GetItem("ProdNo").Specific as EditText;
            lblName = GetItem("ProdDesc").Specific as StaticText;
            cboVersion = GetItem("Item_11").Specific as ComboBox;
            btnSave = GetItem("Item_9").Specific as Button;
            btnClose = GetItem("Item_10").Specific as Button;

            // Init services
            matrixService = new MatrixService(UIAPIRawForm, matrix);
            modeService = new ModeHandlerService(UIAPIRawForm, matrixService, edtItem, lblName, cboVersion, btnSave);
            cflService = new CFLService(UIAPIRawForm, matrixService);
            menuService = new MenuService(UIAPIRawForm.UniqueID);
            dataService = new DataService(UIAPIRawForm, matrix, edtItem, cboVersion);
            navService = new NavigationService();

            // Set menu delete
            menuService.AddMenuItem("1280", "KER_DELETE_ROW", "Delete Row");

            // Initial mode
            modeService.SetAddMode();

            menuService.EnableNavigation(true);

            dataService.InitComboBoxVersion();
            // Tombol
            btnSave.ClickBefore += BtnSave_ClickBefore;
            btnClose.ClickBefore += BtnClose_ClickBefore;
        }

        // ======================================================================
        // BUTTON EVENTS
        // ======================================================================

        private void BtnClose_ClickBefore(object s, SBOItemEventArg e, out bool BubbleEvent)
        {
            BubbleEvent = true;

            Program.SBO_Application.ItemEvent -= OnItemEvent;
            Program.SBO_Application.MenuEvent -= OnMenuEvent;

            UIAPIRawForm.Close();
        }

        private void BtnSave_ClickBefore(object s, SBOItemEventArg e, out bool BubbleEvent)
        {
            BubbleEvent = true;

            switch (btnSave.Caption)
            {
                case "Add":
                    {
                        string msg;
                        bool ok = dataService.Save(out msg);

                        // Tampilkan pesan hasil operasi
                        Program.SBO_Application.MessageBox(msg);

                        if (ok)
                        {
                            // Unsubscribe event global dan event service dengan aman
                            try
                            {
                                Program.SBO_Application.ItemEvent -= OnItemEvent;
                                Program.SBO_Application.MenuEvent -= OnMenuEvent;
                            }
                            catch { /* ignore */ }                           

                            // Pastikan form masih valid sebelum menutup
                            try
                            {
                                var f = Program.SBO_Application.Forms.Item(UIAPIRawForm.UniqueID);
                                if (f != null)
                                    UIAPIRawForm.Close();
                            }
                            catch
                            {
                                // form sudah tidak valid / sudah ditutup, ignore
                            }
                        }

                        BubbleEvent = false;
                        break;
                    }

                case "Update":
                    dataService.Update(edtItem.Value.Trim());
                    BubbleEvent = false;
                    break;

                case "Find":
                    string code = dataService.FindLastByItemCode(edtItem.Value);
                    if (!string.IsNullOrEmpty(code))
                    {
                        dataService.LoadHeader(code);
                        dataService.LoadDetail(code);
                        modeService.SetOKMode();
                    }
                    BubbleEvent = false;
                    break;
            }
        }

        // ======================================================================
        // MENU EVENTS
        // ======================================================================
     
        private void OnMenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.BeforeAction) return;

            switch (pVal.MenuUID)
            {
                case "KER_DELETE_ROW":
                    int row = matrix.GetNextSelectedRow(0, BoOrderType.ot_RowOrder);
                    matrixService.DeleteRow(row);
                    break;

                case "1281":     // FIND MODE
                    modeService.SetFindMode();
                    break;

                case "1282":     // ADD MODE
                    modeService.SetAddMode();
                    break;

                case "1288":     // FIRST
                    navService.LoadKeys();
                    string first = navService.First();
                    if (first != null)
                    {
                        dataService.LoadHeader(first);
                        dataService.LoadDetail(first);
                        modeService.SetOKMode();
                    }
                    break;

                case "1289":     // PREVIOUS
                    string prev = navService.Previous();
                    if (prev != null)
                    {
                        dataService.LoadHeader(prev);
                        dataService.LoadDetail(prev);
                        modeService.SetOKMode();
                    }
                    break;

                case "1290":     // NEXT
                    string next = navService.Next();
                    if (next != null)
                    {
                        dataService.LoadHeader(next);
                        dataService.LoadDetail(next);
                        modeService.SetOKMode();
                    }
                    break;

                case "1291":     // LAST
                    string last = navService.Last();
                    if (last != null)
                    {
                        dataService.LoadHeader(last);
                        dataService.LoadDetail(last);
                        modeService.SetOKMode();
                    }
                    break;
            }
        }

        // ======================================================================
        // ITEM EVENTS (CFL / Matrix click)
        // ======================================================================

        private void OnItemEvent(string formUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (formUID != UIAPIRawForm.UniqueID) return;

            // Simpan posisi klik matrix untuk CFL
            if (pVal.EventType == BoEventTypes.et_CLICK && pVal.BeforeAction)
            {
                lastRow = pVal.Row;
                lastCol = pVal.ColUID;
            }

            // CFL DIPILIH
            if (pVal.EventType == BoEventTypes.et_CHOOSE_FROM_LIST && !pVal.BeforeAction)
            {
                var (code, name) = cflService.ExtractSelected(pVal);

                if (pVal.ItemUID == "ProdNo")     // CFL header
                {
                    var result = cflService.HandleHeaderCFL(code, name);

                    edtItem.Value = result.ItemCode;
                    lblName.Caption = result.ItemName;
                    return;
                }
                else if (pVal.ItemUID == matrix.Item.UniqueID)   // CFL matrix
                {
                    cflService.HandleMatrixCFL(lastRow, lastCol, code, name);
                    return;
                }
            }
        }
    }
}
