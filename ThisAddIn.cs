using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Tools = Microsoft.Office.Tools;

namespace DicePowerPointAddIn
{
  
    public partial class ThisAddIn
    {
        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
           
            Tools.Ribbon.RibbonDropDownItem[] item = new Tools.Ribbon.RibbonDropDownItem[3];
            List<string> ItemNames = new List<string> { "None","German", "English"};
            for (int i = 0; i <= 2; i++)
            {
                item[i] = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item[i].Label = ItemNames[i];
                Globals.Ribbons.Ribbon1.CL_DropDown1.Items.Add(item[i]);
            }
            
        }
        
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
