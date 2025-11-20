using System;
using System.Linq;
using SAPbobsCOM;
using SAPbouiCOM;
using BOM_Version.Helpers;
using System.Threading.Tasks;

namespace BOM_Version
{
    class Menu
    {
        private Application SBO_Application;

        public Menu(Application sboApp)
        {
            SBO_Application = sboApp;
            SBO_Application.MenuEvent += SBO_Application_MenuEvent;
            SBO_Application.ItemEvent += SBO_Application_ItemEvent;
        }

        public void AddMenuItems()
        {
            try
            {
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
        {
            BubbleEvent = true;

            try
            {
                if (!pVal.BeforeAction && pVal.MenuUID == "BOM_Version.Form1")
                {
                    Form1 activeForm = new Form1();
                    activeForm.Show();
                }
            }
            catch (Exception ex)
            {
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
            }
        }



        private void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                
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
                }
            }
            catch (Exception ex)
            {
                Program.SBO_Application.StatusBar.SetText(
                    "ItemEvent Error: " + ex.Message,
                    BoMessageTime.bmt_Short,
                    BoStatusBarMessageType.smt_Error);
            }
        }



    }
}