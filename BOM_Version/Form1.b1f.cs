using System;
using System.Collections.Generic;
<<<<<<< HEAD
=======
using System.Xml;
>>>>>>> 2a6982bf30543cf842f9577408eda331cdfd2d2e
using SAPbouiCOM.Framework;
using BOM_Version.Helpers;

namespace BOM_Version
{
    [FormAttribute("BOM_Version.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
<<<<<<< HEAD
        //private SAPbouiCOM.StaticText lblVersion;
        private SAPbouiCOM.Matrix matrixBOM;

        private SAPbouiCOM.Application SBO_Application = Application.SBO_Application;
        private SAPbobsCOM.Company oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();

        private SAPbouiCOM.DataTable dtBOM;

        public string SelectedItemCode { get; set; }
        public string CallerFormUID { get; set; }

        private Dictionary<string, List<(string ItemCode, string Qty, string Unit, string WareH, string ItemName, string iIssueM)>> versionData =
            new Dictionary<string, List<(string, string, string, string, string, string)>>();

        private Dictionary<string, bool> expandState = new Dictionary<string, bool>();

        public Form1()
        {
            SBO_Application.ItemEvent += SBO_Application_ItemEvent;
        }

        public override void OnInitializeComponent()
        {
            //lblVersion = (SAPbouiCOM.StaticText)this.GetItem("Item_0").Specific;
            matrixBOM = (SAPbouiCOM.Matrix)this.GetItem("Item_3").Specific;

            System.Threading.Tasks.Task.Delay(20).ContinueWith(_ =>
            {
                OnCustomInitialize();
            });

        }

        public override void OnInitializeFormEvents() { }

        private void OnCustomInitialize()
        {

            SAPbouiCOM.Item btn = UIAPIRawForm.Items.Add("btnTest", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            btn.Left = 16;
            btn.Top = 210;
            btn.Width = 80;
            btn.Height = 20;

            ((SAPbouiCOM.Button)btn.Specific).Caption = "Choose";



            // Buat atau ambil DataTable
            try
            {
                dtBOM = UIAPIRawForm.DataSources.DataTables.Item("DT_BOM");
                while (dtBOM.Columns.Count > 0)
                    dtBOM.Columns.Remove(dtBOM.Columns.Item(0).Name);
            }
            catch
            {
                dtBOM = UIAPIRawForm.DataSources.DataTables.Add("DT_BOM");
            }

            // Tambah kolom            
            dtBOM.Columns.Add("Version", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
            dtBOM.Columns.Add("DetCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
            dtBOM.Columns.Add("Qty", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
            dtBOM.Columns.Add("Tres", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 5);
            dtBOM.Columns.Add("Unit", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
            dtBOM.Columns.Add("Warehouse", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
            dtBOM.Columns.Add("DetName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
            dtBOM.Columns.Add("Select", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
            dtBOM.Columns.Add("IssueM", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);

            // Bind Matrix 

            matrixBOM.Columns.Item("colVersion").DataBind.Bind("DT_BOM", "Version");
            matrixBOM.Columns.Item("colDetCode").DataBind.Bind("DT_BOM", "DetCode");
            matrixBOM.Columns.Item("colQty").DataBind.Bind("DT_BOM", "Qty");
            matrixBOM.Columns.Item("colTres").DataBind.Bind("DT_BOM", "Tres");
            matrixBOM.Columns.Item("colUom").DataBind.Bind("DT_BOM", "Unit");
            matrixBOM.Columns.Item("colWare").DataBind.Bind("DT_BOM", "Warehouse");
            matrixBOM.Columns.Item("colNames").DataBind.Bind("DT_BOM", "DetName");
            matrixBOM.Columns.Item("colPilih").DataBind.Bind("DT_BOM", "Select");
            matrixBOM.Columns.Item("colIssue").DataBind.Bind("DT_BOM", "IssueM");
            
            LoadData();
            RebuildMatrix();
        }

        #region Load & Build Matrix

        private void LoadData()
        {
            try
            {


                var oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                // Gunakan interpolated string untuk memasukkan SelectedItemCode
                string sql = $@"
            SELECT H.U_Version, H.U_BOMName, L.U_ItemCode, L.U_Quantity, L.U_Unit, L.U_Warehouse, L.U_ItemName, L.U_IssueM
            FROM [@BOM_VERSION] H
            INNER JOIN [@BOM_VERSION_L] L ON L.Code = H.Code
             WHERE H.U_BOMName = '{SelectedItemCode}'
            ORDER BY H.U_BOMName, H.U_Version";

                oRS.DoQuery(sql);

                versionData.Clear();
                expandState.Clear();

                while (!oRS.EoF)
                {
                    string version = oRS.Fields.Item("U_Version").Value.ToString();
                    string itemCode = oRS.Fields.Item("U_ItemCode").Value.ToString();
                    string qty = oRS.Fields.Item("U_Quantity").Value.ToString();
                    string unit = oRS.Fields.Item("U_Unit").Value.ToString();
                    string whs = oRS.Fields.Item("U_Warehouse").Value.ToString();
                    string iName = oRS.Fields.Item("U_ItemName").Value.ToString();
                    string iIssueM = oRS.Fields.Item("U_IssueM").Value.ToString();

                    if (!versionData.ContainsKey(version))
                    {
                        versionData[version] = new List<(string, string, string, string, string, string)>();
                        expandState[version] = false;
                    }

                    versionData[version].Add((itemCode, qty, unit, whs, iName, iIssueM));

                    oRS.MoveNext();
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Error LoadData: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }



        private void RebuildMatrix()
        {
            try
            {
                UIAPIRawForm.Freeze(true);
                dtBOM.Rows.Clear();

                // Loop per Version
                foreach (var ver in versionData)
                {
                    string version = ver.Key;

                    // HEADER VERSION
                    dtBOM.Rows.Add();
                    int header = dtBOM.Rows.Count - 1;

                    dtBOM.SetValue("Version", header, version);
                    dtBOM.SetValue("Tres", header, expandState[version] ? "▼" : "►");
                    dtBOM.SetValue("DetCode", header, "");
                    dtBOM.SetValue("Qty", header, "");
                    dtBOM.SetValue("Unit", header, "");
                    dtBOM.SetValue("Warehouse", header, "");
                    dtBOM.SetValue("DetName", header, "");
                    dtBOM.SetValue("IssueM", header, "");

                    // DETAIL COMPONENT
                    if (expandState[version])
                    {
                        foreach (var comp in ver.Value)
                        {
                            dtBOM.Rows.Add();
                            int row = dtBOM.Rows.Count - 1;

                            dtBOM.SetValue("Version", row, ""); // kosong
                            dtBOM.SetValue("Tres", row, "");
                            dtBOM.SetValue("DetCode", row, comp.ItemCode);
                            dtBOM.SetValue("Qty", row, comp.Qty);
                            dtBOM.SetValue("Unit", row, comp.Unit);
                            dtBOM.SetValue("Warehouse", row, comp.WareH);
                            dtBOM.SetValue("DetName", row, comp.ItemName);
                            dtBOM.SetValue("IssueM", row, comp.iIssueM);
                        }
                    }
                }

                matrixBOM.LoadFromDataSource();
                matrixBOM.AutoResizeColumns();



                matrixBOM.LoadFromDataSource();
                matrixBOM.AutoResizeColumns();
            }
            finally
            {
                UIAPIRawForm.Freeze(false);
            }
        }

        public void ToggleExpand(int row)
        {
            string version = dtBOM.GetValue("Version", row).ToString();
            if (string.IsNullOrWhiteSpace(version)) return;

            expandState[version] = !expandState[version];
            RebuildMatrix();
        }


        #endregion

        #region Event Handlers
        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (FormUID != UIAPIRawForm.UniqueID)
                return;

            // Expand / Collapse
            if (pVal.ItemUID == "Item_3" &&
                pVal.ColUID == "colTres" &&
                pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK &&
                !pVal.BeforeAction)
            {
                ToggleExpand(pVal.Row - 1);
                return;
            }

            // HANDLE CHECKBOX SINGLE SELECTION
            if (pVal.ItemUID == "Item_3" &&
                pVal.ColUID == "colPilih" &&
                pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK &&
                !pVal.BeforeAction)
            {
                HandleSingleCheckbox(pVal.Row - 1);
                return;
            }

            if (pVal.ItemUID == "btnTest" &&
            pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK &&
            !pVal.BeforeAction)
            {
                string version = GetSelectedVersion();
                ApplyBOMVersion(version);
                return;
            }

            
        }
        #endregion
        private void HandleSingleCheckbox(int row)
        {
            try
            {
                UIAPIRawForm.Freeze(true);

                // Jika row tidak valid, keluar
                if (row < 0 || row >= dtBOM.Rows.Count)
                    return;

                // Toggle baris yang diklik menjadi Y
                dtBOM.SetValue("Select", row, "Y");

                // Semua baris lain harus menjadi N
                for (int i = 0; i < dtBOM.Rows.Count; i++)
                {
                    if (i != row)
                        dtBOM.SetValue("Select", i, "N");
                }

                matrixBOM.LoadFromDataSource();

                // 🔥 Highlight row yang dipilih
                HighlightSelectedRow(row);
            }
            finally
            {
                UIAPIRawForm.Freeze(false);
            }
        }

        private void HighlightSelectedRow(int selectedRow)
        {
            // Warna highlight (kuning lembut)
            int highlight = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow);

            // Reset semua baris
            for (int r = 0; r < dtBOM.Rows.Count; r++)
            {
                matrixBOM.CommonSetting.SetRowBackColor(r + 1, -1);
                // reset ke default
            }

            // Set warna highlight untuk row yg dipilih
            matrixBOM.CommonSetting.SetRowBackColor(selectedRow + 1, highlight);
        }

        private string GetSelectedVersion()
        {
            for (int i = 0; i < dtBOM.Rows.Count; i++)
            {
                string sel = dtBOM.GetValue("Select", i).ToString();
                string ver = dtBOM.GetValue("Version", i).ToString();

                // hanya baris header memiliki Version
                if (sel == "Y" && !string.IsNullOrWhiteSpace(ver))
                    return ver;
            }

            return null;
        }

        private void ApplyBOMVersion(string version)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(version))
                {
                    SBO_Application.MessageBox("Tidak ada Version yang dipilih.");
                    return;
                }

                int ask = SBO_Application.MessageBox(
                    $"Yakin ganti BOM Version menjadi '{version}' ?",
                    1, "Ya", "Tidak"
                );
                if (ask != 1) return;

                string targetFormUID = CallerFormUID;
                if (string.IsNullOrEmpty(targetFormUID))
                {
                    foreach (SAPbouiCOM.Form f in SBO_Application.Forms)
                        if (f.TypeEx == "65211")
                            targetFormUID = f.UniqueID;
                }

                if (string.IsNullOrEmpty(targetFormUID))
                {
                    SBO_Application.MessageBox("Form Production Order tidak ditemukan.");
                    return;
                }

                BomHelper.UpdateTextBom(
                    oCompany,
                    targetFormUID,
                    version,                  
                    null
                );

                SAPbouiCOM.Form oProdForm = SBO_Application.Forms.Item(targetFormUID);
                string docNum = ((SAPbouiCOM.EditText)oProdForm.Items.Item("18").Specific).Value;

                BomHelper.UpdateByDB(
                    oCompany,
                    targetFormUID,
                    version,
                    docNum
                );

                SBO_Application.StatusBar.SetText(
                    $"BOM Version '{version}' berhasil diterapkan & Production Order ter-update.",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Success
                );

                UIAPIRawForm.Close();

                BomHelper.RefreshProductionOrderUI(targetFormUID, docNum);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox("Error: " + ex.Message);
            }
        }


    }
}
=======
        public Form1()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            

        }

        private void OnCustomInitialize()
        {
        }
        
    }
}
>>>>>>> 2a6982bf30543cf842f9577408eda331cdfd2d2e
