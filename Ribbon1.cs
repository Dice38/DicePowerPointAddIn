using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;



namespace DicePowerPointAddIn
{
    public partial class Ribbon1
    {
        
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            var Slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;
            switch (this.CL_DropDown1.SelectedItem.Label)
            {
                case "None":
                    break;
                case "German":
                    if (Slides.Count > 0)
                    {
                        for (int i = 1; i <= Slides.Count; i++)
                        {
                            int j = 1;
                            while (j <= Slides[i].Shapes.Count)
                            {
                                Slides[i].Shapes[j].TextFrame.TextRange.LanguageID = Office.MsoLanguageID.msoLanguageIDGerman;
                                j++;

                            }
                        }

                    }
                    break;
                case "English":
                    if (Slides.Count > 0)
                    {
                        for (int i = 1; i <= Slides.Count; i++)
                        {
                            int j = 1;
                            while (j <= Slides[i].Shapes.Count)
                            {
                                Slides[i].Shapes[j].TextFrame.TextRange.LanguageID = Office.MsoLanguageID.msoLanguageIDEnglishUK;
                                j++;

                            }
                        }


                    }
                    break;
                default:
                    break;
            }

        }

        private void CL_DropDown1_SelectionChanged(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
