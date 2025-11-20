using SAPbouiCOM.Framework;
using SAPbobsCOM;
using System;

namespace BOM_Version.Services
{
    public class ErrorHandlerService
    {
        // ============================================================
        // DI API ERROR FORMATTER
        // ============================================================
        public string GetDIError(Company company)
        {
            company.GetLastError(out int errCode, out string errMsg);

            if (errCode == 0)
                return "Unknown DI Error";

            return $"[DI-{errCode}] {errMsg}";
        }

        // ============================================================
        // MESSAGE POPUP
        // ============================================================
        public void ShowError(string msg)
        {
            Application.SBO_Application.MessageBox("Error: " + msg);
        }

        public void ShowWarning(string msg)
        {
            Application.SBO_Application.StatusBar.SetText(
                msg,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Warning
            );
        }

        public void ShowSuccess(string msg)
        {
            Application.SBO_Application.StatusBar.SetText(
                msg,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Success
            );
        }

        // ============================================================
        // SAFE EXECUTION WRAPPER
        // ============================================================
        public void TryDo(Action action, string context = "")
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                ShowError($"{context} - {ex.Message}");
            }
        }
    }
}
