using System;
using SAPbouiCOM.Framework;
using BOM_Version.Services;

namespace BOM_Version
{
    class Program
    {
        public static SAPbouiCOM.Application SBO_Application;

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application oApp = null;

                if (args.Length < 1)
                {
                    // Connect ke SAP B1
                    oApp = new Application();
                    SBO_Application = Application.SBO_Application;
                }
                else
                {
                    // Jika ada koneksi string dari Add-on
                    oApp = new Application(args[0]);
                    SBO_Application = Application.SBO_Application;
                }

                // Inisialisasi Menu dan register event
                Menu myMenu = new Menu(SBO_Application);
                myMenu.AddMenuItems();

                // Register AppEvent (Shutdown, CompanyChanged, dll)
                SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);

                // Jalankan Add-on
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error Add-on: " + ex.Message);
            }
        }

        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    CompanyService.Disconnect();
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    // Optional handling
                    break;
                default:
                    break;
            }
        }
    }
}
