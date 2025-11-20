using System;
using System.IO;
using SAPbouiCOM;

namespace BOM_Version
{
    class Menu
    {
        private readonly Application SBO_Application;

        public Menu(Application sboApp)
        {
            SBO_Application = sboApp;
            SBO_Application.ItemEvent += SBO_Application_ItemEvent;
            SBO_Application.MenuEvent += SBO_Application_MenuEvent;

            AddMenuItems();
        }

        public void AddMenuItems()
        {
            try
            {
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
        {
            BubbleEvent = true;

            try
            {
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
                }
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox("MenuEvent Error: " + ex.Message);
            }
        }


        // ----------------------------------------------------
        // ITEM EVENT HANDLER
        // ----------------------------------------------------

        private void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
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
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(
                    "ItemEvent Error: " + ex.Message,
                    BoMessageTime.bmt_Short,
                    BoStatusBarMessageType.smt_Error);
            }
        }


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



