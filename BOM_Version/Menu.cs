using System;
<<<<<<< HEAD
using System.IO;
using SAPbouiCOM;
=======
using System.Linq;
using SAPbobsCOM;
using SAPbouiCOM;
using BOM_Version.Helpers;
using System.Threading.Tasks;
>>>>>>> 2a6982bf30543cf842f9577408eda331cdfd2d2e

namespace BOM_Version
{
    class Menu
    {
<<<<<<< HEAD
        private readonly Application SBO_Application;
=======
        private Application SBO_Application;
>>>>>>> 2a6982bf30543cf842f9577408eda331cdfd2d2e

        public Menu(Application sboApp)
        {
            SBO_Application = sboApp;
<<<<<<< HEAD
            SBO_Application.ItemEvent += SBO_Application_ItemEvent;
            SBO_Application.MenuEvent += SBO_Application_MenuEvent;

            AddMenuItems();
=======
            SBO_Application.MenuEvent += SBO_Application_MenuEvent;
            SBO_Application.ItemEvent += SBO_Application_ItemEvent;
>>>>>>> 2a6982bf30543cf842f9577408eda331cdfd2d2e
        }

        public void AddMenuItems()
        {
            try
            {
<<<<<<< HEAD
                // Ambil root menu SAP
                SAPbouiCOM.Menus rootMenus = SBO_Application.Menus;

                // 1️⃣ Cari menu Production (ID default = "43520")
                SAPbouiCOM.MenuItem productionMenu = null;
                try
                {
                    productionMenu = rootMenus.Item("4352");
                }
                catch
                {
                    SBO_Application.MessageBox("Menu Production tidak ditemukan.");
                    return;
                }

                // 2️⃣ Cari menu Bill of Materials (ID default = "43616")
                SAPbouiCOM.MenuItem bomMenu;
                try
                {
                    bomMenu = productionMenu.SubMenus.Item("4353");
                }
                catch
                {
                    SBO_Application.MessageBox("Menu BOM tidak ditemukan.");
                    return;
                }

                // 3️⃣ Jika menu sudah ada, jangan buat lagi
                foreach (SAPbouiCOM.MenuItem m in productionMenu.SubMenus)
                {
                    if (m.UID == "ALT_BOM")
                        return; // sudah ada
                }

                // 4️⃣ Cari posisi Bill of Materials pada parent list
                int bomPosition = -1;
                for (int i = 0; i < productionMenu.SubMenus.Count; i++)
                {
                    SAPbouiCOM.MenuItem item = productionMenu.SubMenus.Item(i);
                    if (item.UID == "4353") // Bill of Materials
                    {
                        bomPosition = i;
                        break;
                    }
                }

                if (bomPosition == -1)
                    bomPosition = productionMenu.SubMenus.Count - 1;

                // 5️⃣ Buat menu baru tepat setelah Bill of Materials
                SAPbouiCOM.MenuCreationParams cp =
                    (SAPbouiCOM.MenuCreationParams)SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);

                cp.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                cp.UniqueID = "ALT_BOM";
                cp.String = "Alternate Bill Of Material";
                cp.Enabled = true;

                cp.Position = bomPosition + 1;

                productionMenu.SubMenus.AddEx(cp);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox("Error AddMenuItems: " + ex.Message);
            }
        }

        private SAPbouiCOM.Form GetForm1()
        {
            foreach (SAPbouiCOM.Form f in SBO_Application.Forms)
            {
                if (f.TypeEx == "BOM_Version.Form1")
                    return f;
            }
            return null;
        }



        private void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
=======
                Menus oMenus = SBO_Application.Menus;
                MenuItem oMenuItem = oMenus.Item("43520"); // Modules

                MenuCreationParams oCreationPackage = (MenuCreationParams)SBO_Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Type = BoMenuType.mt_POPUP;
                oCreationPackage.UniqueID = "BOM_Version";
                oCreationPackage.String = "BOM Version";
                oCreationPackage.Enabled = true;
                oCreationPackage.Position = -1;

                Menus subMenus = oMenuItem.SubMenus;
                try { subMenus.AddEx(oCreationPackage); } catch { }

                oMenuItem = SBO_Application.Menus.Item("BOM_Version");
                subMenus = oMenuItem.SubMenus;

                oCreationPackage.Type = BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "BOM_Version.Form1";
                oCreationPackage.String = "Form1";

                try { subMenus.AddEx(oCreationPackage); }
                catch { SBO_Application.SetStatusBarMessage("Menu already exists", BoMessageTime.bmt_Short, true); }
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox("Error adding menu: " + ex.Message);
            }
        }

        private void SBO_Application_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
>>>>>>> 2a6982bf30543cf842f9577408eda331cdfd2d2e
        {
            BubbleEvent = true;

            try
            {
<<<<<<< HEAD
                // Cek menu ALT_BOM
                if (!pVal.BeforeAction && pVal.MenuUID == "ALT_BOM")
                {
                    // Jangan buka dua kali
                    foreach (SAPbouiCOM.Form f in SBO_Application.Forms)
                    {
                        if (f.TypeEx == "BOM_Version.Alternate")
                        {
                            f.Select();
                            return;
                        }
                    }

                    Alternate alt = new Alternate();
                    alt.Show();
=======
                if (!pVal.BeforeAction && pVal.MenuUID == "BOM_Version.Form1")
                {
                    Form1 activeForm = new Form1();
                    activeForm.Show();
>>>>>>> 2a6982bf30543cf842f9577408eda331cdfd2d2e
                }
            }
            catch (Exception ex)
            {
<<<<<<< HEAD
                SBO_Application.MessageBox("MenuEvent Error: " + ex.Message);
=======
                SBO_Application.MessageBox("Menu Event Error: " + ex.ToString());
            }
        }

        private void AddLoadBomButton(Form oForm)
        {
            try
            {
                //// clone existing button (misal OK) untuk ukuran & pos
                //Item baseItem = oForm.Items.Item("1"); // OK button
                //Item newBtn = oForm.Items.Add("btnLoadBOM", BoFormItemTypes.it_BUTTON);

                //newBtn.Left = baseItem.Left + 200;
                //newBtn.Top = baseItem.Top;
                //newBtn.Height = baseItem.Height;
                //newBtn.Width = baseItem.Width;

                //((Button)newBtn.Specific).Caption = "Load BOM";

                //// clone existing button (misal OK) untuk ukuran & pos
                ////Item baseItem = oForm.Items.Item("1"); // OK button
                //Item newRef = oForm.Items.Add("btnRefresh", BoFormItemTypes.it_BUTTON);

                //newRef.Left = baseItem.Left + 300;
                //newRef.Top = baseItem.Top;
                //newRef.Height = baseItem.Height;
                //newRef.Width = baseItem.Width;

                //((Button)newRef.Specific).Caption = "Refresh";


                Item BomItem = oForm.Items.Item("540000153");
                // --- TEXTBOX untuk BOM Version ---
                Item txtBomVer = oForm.Items.Add("txtBOMVER", BoFormItemTypes.it_EDIT);

                // Letakkan di atas tombol Refresh
                txtBomVer.Left = BomItem.Left;
                txtBomVer.Top = BomItem.Top + 17;
                txtBomVer.Width = 110;
                txtBomVer.Height = 15;

                // OPTIONAL: label untuk textbox
                Item lblBomVer = oForm.Items.Add("lblBOMVER", BoFormItemTypes.it_STATIC);
                lblBomVer.Left = txtBomVer.Left - 135;
                lblBomVer.Top = txtBomVer.Top;
                lblBomVer.Width = 80;
                lblBomVer.Height = 15;
                ((StaticText)lblBomVer.Specific).Caption = "BOM Version :";

                // Bind textbox ke UDF U_BOMVER
                ((EditText)txtBomVer.Specific).DataBind.SetBound(true, "OWOR", "U_BOMVER");


            }
            catch (Exception ex)
            {
                Program.SBO_Application.StatusBar.SetText(
                    "Error add button: " + ex.Message,
                    BoMessageTime.bmt_Short,
                    BoStatusBarMessageType.smt_Error
                );
            }
        }

        private void HandleBOMVersionChange(string FormUID)
        {
            Form oForm = Program.SBO_Application.Forms.Item(FormUID);

            string cekStatus =
                ((SAPbouiCOM.ComboBox)oForm.Items.Item("10").Specific).Value?.Trim();
            // Hanya jalan jika status = "P"
            if (cekStatus != "P")
                return;

            string bomVerValue =
                ((SAPbouiCOM.EditText)oForm.Items.Item("txtBOMVER").Specific).Value?.Trim();

            if (string.IsNullOrEmpty(bomVerValue))
                return;

           
            string docNum =
                ((SAPbouiCOM.EditText)oForm.Items.Item("18").Specific).Value?.Trim();

            if (string.IsNullOrEmpty(docNum))
                return;

            var oCompany = BOM_Version.Services.CompanyService.GetCompany();

            BomHelper.UpdateByDB(oCompany, FormUID, bomVerValue, docNum);
        }

        private void HandleButtonPressed(string FormUID, string itemUID)
        {
            Form oForm = Program.SBO_Application.Forms.Item(FormUID);

            switch (itemUID)
            {
                case "btnLoadBOM":
                    Program.SBO_Application.StatusBar.SetText(
                        "BOM Version executed.",
                        BoMessageTime.bmt_Short,
                        BoStatusBarMessageType.smt_Success);
                    break;

                case "btnRefresh":
                    Program.SBO_Application.StatusBar.SetText(
                        "Refresh executed.",
                        BoMessageTime.bmt_Short,
                        BoStatusBarMessageType.smt_Success);
                    break;
>>>>>>> 2a6982bf30543cf842f9577408eda331cdfd2d2e
            }
        }


<<<<<<< HEAD
        // ----------------------------------------------------
        // ITEM EVENT HANDLER
        // ----------------------------------------------------
=======
>>>>>>> 2a6982bf30543cf842f9577408eda331cdfd2d2e

        private void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
<<<<<<< HEAD
                // Add button on form 65211
                if (pVal.EventType == BoEventTypes.et_FORM_LOAD &&
                    pVal.FormTypeEx == "65211" &&
                    !pVal.BeforeAction)
                {
                    AddDebugButton(FormUID);
                }

                // Add button on form 672
                if (pVal.EventType == BoEventTypes.et_FORM_LOAD &&
                    pVal.FormTypeEx == "672" &&
                    !pVal.BeforeAction)
                {
                    AddAltenateBtn(FormUID);
                }

                // Handle button click on form 65211
                if (pVal.ItemUID == "btnDEBUG" &&
                    pVal.FormTypeEx == "65211" &&
                    pVal.EventType == BoEventTypes.et_ITEM_PRESSED &&
                    !pVal.BeforeAction)
                {
                    OpenForm1WithItemCode();
                }

                // Handle button click on form 672
                if (pVal.ItemUID == "btnAlt" &&
                    pVal.FormTypeEx == "672" &&
                    pVal.EventType == BoEventTypes.et_ITEM_PRESSED &&
                    !pVal.BeforeAction)
                {                   
                    Alternate alt = new Alternate();
                    alt.Show();
=======
                
                if (pVal.EventType == BoEventTypes.et_FORM_LOAD
                    && pVal.FormTypeEx == "65211"
                    && !pVal.BeforeAction)
                {
                    Form oForm = Program.SBO_Application.Forms.Item(FormUID);
                    AddLoadBomButton(oForm);
                    return;
                }


                if (pVal.FormTypeEx == "65211"
                    && pVal.EventType == BoEventTypes.et_VALIDATE
                    && !pVal.BeforeAction
                    && pVal.ItemUID == "txtBOMVER")
                {
                    HandleBOMVersionChange(FormUID);
                    return;
                }

                
                if (pVal.FormTypeEx == "65211"
                    && pVal.EventType == BoEventTypes.et_ITEM_PRESSED
                    && !pVal.BeforeAction)
                {
                    //HandleButtonPressed(FormUID, pVal.ItemUID);
                    return;
>>>>>>> 2a6982bf30543cf842f9577408eda331cdfd2d2e
                }
            }
            catch (Exception ex)
            {
<<<<<<< HEAD
                SBO_Application.StatusBar.SetText(
=======
                Program.SBO_Application.StatusBar.SetText(
>>>>>>> 2a6982bf30543cf842f9577408eda331cdfd2d2e
                    "ItemEvent Error: " + ex.Message,
                    BoMessageTime.bmt_Short,
                    BoStatusBarMessageType.smt_Error);
            }
        }


<<<<<<< HEAD
        // ----------------------------------------------------
        // ADD DEBUG BUTTON SAFELY
        // ----------------------------------------------------
        private void AddDebugButton(string formUID)
        {
            Form oForm = SBO_Application.Forms.Item(formUID);

            // Skip if button already exists
            if (ItemExists(oForm, "btnDEBUG"))
                return;

            oForm.Freeze(true);
            try
            {
                Item btn = oForm.Items.Add("btnDEBUG", BoFormItemTypes.it_BUTTON);
                btn.Left = 300;
                btn.Top = 10;
                btn.Width = 110;
                btn.Height = 20;

                ((Button)btn.Specific).Caption = "Alternate BOM";
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        private void AddAltenateBtn(string formUID)
        {
            Form oForm = SBO_Application.Forms.Item(formUID);

            // Skip if button already exists
            if (ItemExists(oForm, "btnAlt"))
                return;

            oForm.Freeze(true);
            try
            {
                Item oItemRef = oForm.Items.Item("2"); // contoh
                int left = oItemRef.Left + 80;
                int top = oItemRef.Top;

                Item btn = oForm.Items.Add("btnAlt", BoFormItemTypes.it_BUTTON);
                btn.Left = left;
                btn.Top = top;
                btn.Width = 90;
                btn.Height = 20;

                ((Button)btn.Specific).Caption = "Alternate BOM";
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        private bool ItemExists(Form oForm, string itemUID)
        {
            try
            {
                var tmp = oForm.Items.Item(itemUID);
                return true;
            }
            catch
            {
                return false;
            }
        }





        // ----------------------------------------------------
        // BUTTON CLICK → OPEN FORM1 & SEND ITEMCODE
        // ----------------------------------------------------      

        private void OpenForm1WithItemCode()
        {
            try
            {
                // 🔍 Cek apakah Form1 sudah ada
                SAPbouiCOM.Form existingForm = GetForm1();
                if (existingForm != null)
                {
                    // 🔥 Jika minimize → restore
                    if (existingForm.State == SAPbouiCOM.BoFormStateEnum.fs_Minimized)
                        existingForm.State = SAPbouiCOM.BoFormStateEnum.fs_Restore;

                    // 🔥 Pastikan form terlihat
                    existingForm.Visible = true;

                    // 🔥 Bring to front / fokus
                    existingForm.Select();

                    return;
                }

                // ============ Form belum ada → buat baru ===========
                // Cari Form Production Order
                SAPbouiCOM.Form prodForm = null;
                string formUID = null;

                foreach (SAPbouiCOM.Form f in SBO_Application.Forms)
                {
                    if (f.TypeEx == "65211")
                    {
                        prodForm = f;
                        formUID = f.UniqueID;
                        break;
                    }
                }

                if (prodForm == null)
                {
                    SBO_Application.MessageBox("Form Production Order tidak ditemukan.");
                    return;
                }

                // Ambil ItemCode
                string itemCode = null;

                try
                {
                    var ds = prodForm.DataSources.DBDataSources.Item("OWOR");
                    itemCode = ds.GetValue("ItemCode", 0).Trim();
                }
                catch { }

                if (string.IsNullOrWhiteSpace(itemCode))
                    itemCode = ((SAPbouiCOM.EditText)prodForm.Items.Item("6").Specific).Value.Trim();

                if (string.IsNullOrWhiteSpace(itemCode))
                {
                    SBO_Application.MessageBox("ItemCode tidak ditemukan.");
                    return;
                }

                // Buat Form1 baru
                Form1 f1 = new Form1
                {
                    CallerFormUID = formUID,
                    SelectedItemCode = itemCode
                };

                f1.Show();
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox("Error OpenForm1: " + ex.Message);
            }
        }



    }


}



=======

    }
}
>>>>>>> 2a6982bf30543cf842f9577408eda331cdfd2d2e
