using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Word;

namespace Replace
{
    public partial class Replace
    {
        private void Zamień_znaki_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonRun_Click(object sender, RibbonControlEventArgs e)
        {
            string[] ArrayOfCharacters = new string[6];
            ArrayOfCharacters[0] = "a";
            ArrayOfCharacters[1] = "o";
            ArrayOfCharacters[2] = "u";
            ArrayOfCharacters[3] = "w";
            ArrayOfCharacters[4] = "i";
            ArrayOfCharacters[5] = "z";

            Microsoft.Office.Interop.Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            foreach(string character in ArrayOfCharacters)
            {
                doc.Content.Find.Execute(character + " ", false, true, false, false, false, true, 1, false, character + "\u00A0", 2, false, false, false, false);
            }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            string[] ArrayOfCharacters = new string[6];
            ArrayOfCharacters[0] = "a";
            ArrayOfCharacters[1] = "o";
            ArrayOfCharacters[2] = "u";
            ArrayOfCharacters[3] = "w";
            ArrayOfCharacters[4] = "i";
            ArrayOfCharacters[5] = "z";

            Microsoft.Office.Interop.Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            foreach (string character in ArrayOfCharacters)
            {
                doc.Content.Find.Execute(character + "\u00A0", false, true, false, false, false, true, 1, false, character + " ", 2, false, false, false, false);
            }
        }
    }
}
