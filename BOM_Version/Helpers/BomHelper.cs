using System;
using SAPbobsCOM;
using SAPbouiCOM.Framework;
using BOM_Version.Services;
using BOM_Version.Model;
using System.Collections.Generic;
using System.Globalization;
using SAPbouiCOM;
using System.Threading.Tasks;

namespace BOM_Version.Helpers
{

    public static class BomHelper
    {

        private static bool isRefreshing = false;

        
        public static void RefreshProductionOrderUI(string formUID, string docNum)
        {
            if (isRefreshing) return;
            isRefreshing = true;

            SAPbouiCOM.Form oForm = Program.SBO_Application.Forms.Item(formUID);

            try
            {
                System.Threading.Thread.Sleep(100);
                oForm.Freeze(true);

                //// Reset perubahan agar SAP B1 tidak memunculkan save confirmation
                if (oForm.Mode == BoFormMode.fm_UPDATE_MODE || oForm.Mode == BoFormMode.fm_ADD_MODE)
                    oForm.Mode = BoFormMode.fm_OK_MODE;

                //// Masuk Find Mode
                oForm.Mode = BoFormMode.fm_FIND_MODE;

                // Set DocNum
                if (oForm.Items.Item("18").Specific is SAPbouiCOM.EditText txtDocNum)
                    txtDocNum.Value = docNum;

                // Klik Find
                oForm.Items.Item("1").Click();

                System.Threading.Thread.Sleep(50); // redraw UI
            }
            catch (Exception ex)
            {
                Program.SBO_Application.StatusBar.SetText(
                    $"Refresh UI Error: {ex.Message}",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm?.Freeze(false);
                isRefreshing = false;
            }
        }

        public static void UpdateTextBom(
    SAPbobsCOM.Company oCompany,
    string formUID,
    string textSel,   
    string formToCloseUID = null)
        {
            SAPbouiCOM.Form oForm = null;

            try
            {
                oForm = Program.SBO_Application.Forms.Item(formUID);
                oForm.Freeze(true);

                // Update EditText
                try
                {
                    var txt = (SAPbouiCOM.EditText)oForm.Items.Item("txtBOMVER").Specific;
                    txt.Value = textSel;
                }
                catch { }

                Program.SBO_Application.StatusBar.SetText(
                    "BOMVer berhasil diset.",
                    BoMessageTime.bmt_Short,
                    BoStatusBarMessageType.smt_Success);

                System.Windows.Forms.Timer t = new System.Windows.Forms.Timer();
                t.Interval = 50; // cukup delay sebentar

                Program.SBO_Application.Forms.Item(formToCloseUID).Close();

                
            }
            catch (Exception ex)
            {
                //Program.SBO_Application.StatusBar.SetText(
                //    "Error UpdateTextBom: " + ex.Message,
                //    BoMessageTime.bmt_Short,
                //    BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (oForm != null)
                    oForm.Freeze(false);
            }
        }


        public static void UpdateByDB(SAPbobsCOM.Company oCompany, string formUID, string bomVerValue, string docNum)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(docNum))
                {
                    Program.SBO_Application.MessageBox("DocNum tidak boleh kosong.");
                    return;
                }

                if (string.IsNullOrWhiteSpace(bomVerValue))
                {
                    Program.SBO_Application.MessageBox("BOM Version tidak boleh kosong.");
                    return;
                }

                // Sanitize nilai input (hindari SQL Injection)
                string safeDocNum = docNum.Replace("'", "");
                string safeBomVer = bomVerValue.Replace("'", "");


                SAPbobsCOM.Recordset oRec =
                    (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string query = $@"
            SELECT DocEntry, ItemCode 
            FROM OWOR 
            WHERE DocNum = {safeDocNum}";

                oRec.DoQuery(query);

                if (oRec.RecordCount == 0)
                {
                    Program.SBO_Application.MessageBox($"Production Order dengan DocNum {docNum} tidak ditemukan.");
                    return;
                }

                int docEntry = Convert.ToInt32(oRec.Fields.Item("DocEntry").Value);
                string headerItemCode = (oRec.Fields.Item("ItemCode").Value ?? "").ToString();

                if (string.IsNullOrWhiteSpace(headerItemCode))
                {
                    Program.SBO_Application.MessageBox("ItemCode pada OWOR kosong. Tidak dapat melanjutkan.");
                    return;
                }


                SAPbobsCOM.ProductionOrders oProd =
                    (SAPbobsCOM.ProductionOrders)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders);

                if (!oProd.GetByKey(docEntry))
                {
                    Program.SBO_Application.MessageBox($"Gagal load Production Order dengan DocEntry {docEntry}");
                    return;
                }


                for (int i = oProd.Lines.Count - 1; i >= 0; i--)
                {
                    oProd.Lines.SetCurrentLine(i);
                    oProd.Lines.Delete();
                }


                SAPbobsCOM.Recordset oRS =
                    (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string queryBom = $@"
            SELECT 
                t1.[U_ItemCode],
                t1.[U_Quantity],
                t1.[U_Warehouse]
            FROM [USER_TRAINING].[dbo].[@BOM_VERSION] t0
            INNER JOIN [@BOM_VERSION_L] t1 ON t0.Code = t1.Code
            WHERE t0.[U_BOMName] = '{headerItemCode.Replace("'", "")}'
              AND t0.[U_Version] = '{safeBomVer}'
            ORDER BY t1.[LineID]";

                oRS.DoQuery(queryBom);

                if (oRS.RecordCount == 0)
                {
                    Program.SBO_Application.MessageBox(
                        $"Tidak ada BOM Version '{bomVerValue}' untuk item '{headerItemCode}'.");
                    return;
                }


                int lineIndex = 0;

                oProd.UserFields.Fields.Item("U_BOMVER").Value = bomVerValue;

                while (!oRS.EoF)
                {
                    string itemCode = oRS.Fields.Item("U_ItemCode").Value.ToString();
                    double qty = Convert.ToDouble(oRS.Fields.Item("U_Quantity").Value);
                    string whs = oRS.Fields.Item("U_Warehouse").Value.ToString();

                    if (string.IsNullOrWhiteSpace(itemCode))
                    {
                        Program.SBO_Application.MessageBox($"Line {lineIndex}: ItemCode kosong, dilewati.");
                        oRS.MoveNext();
                        continue;
                    }

                    oProd.Lines.SetCurrentLine(lineIndex);
                    oProd.Lines.ItemNo = itemCode;
                    oProd.Lines.BaseQuantity = qty;
                    oProd.Lines.Warehouse = whs;

                    oRS.MoveNext();

                    if (!oRS.EoF)
                        oProd.Lines.Add();

                    lineIndex++;
                }


                int updateResult = oProd.Update();

                if (updateResult == 0)
                {
                    //Program.SBO_Application.MessageBox("Production Order berhasil diperbarui dengan BOM Version!");
                }
                else
                {
                    oCompany.GetLastError(out int errCode, out string errMsg);
                    Program.SBO_Application.MessageBox($"Gagal update Production Order: [{errCode}] {errMsg}");
                }
                //RefreshProductionOrderUI(formUID, docNum);
            }
            catch (Exception ex)
            {
                Program.SBO_Application.StatusBar.SetText(
                    $"Error UpdateByDB: {ex.Message}",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
    }
}

