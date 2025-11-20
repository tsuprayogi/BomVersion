using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;
using BOM_Version.Helpers;

namespace BOM_Version
{
    [FormAttribute("BOM_Version.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
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