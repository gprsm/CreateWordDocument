using System;
using System.Collections.Generic;
using System.IO;
using CreateWordDocument.Models;
using Microsoft.Office.Interop.Word;

namespace CreateWordDocument.Helper
{
    public class WordClass
    {
        public void ProcessInputs(List<ExcelModel> excelInputs, string templatePath, string textWins,
            string textPar)
        {
            if (!File.Exists(templatePath))
            {
                var info = new CollectInfo();
                templatePath=info.InteractWithUser("template not Found. enter full word file path again(example 'c:/template.docx')...");
            }
            object oMissing = System.Reflection.Missing.Value;
            Application fileOpen = new Application();
            Document document = fileOpen.Documents.Open(templatePath,ReadOnly:false);
            //Make the file visible 
            fileOpen.Visible = false;
            foreach (var excelInput in excelInputs)
            {
                
            }
        }
        public void FindAndReplace(string path,string textToR,string replace,string resPath)
        {
            
           
            object oMissing = System.Reflection.Missing.Value;
            Application fileOpen = new Application();   
            //Open a already existing word file into the new document created
            Document document = fileOpen.Documents.Open(path,ReadOnly:false);
            //Make the file visible 
            fileOpen.Visible = false;
            //document.Activate();
            //The FindAndReplace takes the text to find under any formatting and replaces it with the
            //new text with the same exact formmating (e.g red bold text will be replaced with red bold text)
            //FindAndReplace(fileOpen, textToR, replace);
            //SearchReplace(fileOpen, textToR, replace);
            SearchTextBox(fileOpen, textToR, replace);
            //Save the editted file in a specified location
            //Can use SaveAs instead of SaveAs2 and just give it a name to have it saved by default
            //to the documents folder
            document.SaveAs2(resPath);
            object outputFileName = resPath.Replace(".docx", ".pdf");
            var pdfPath =outputFileName;
            object fileFormat = WdSaveFormat.wdFormatPDF;
            document.SaveAs( pdfPath,
                ref fileFormat, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            
            object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
            ((_Document)document).Close(ref saveChanges, ref oMissing, ref oMissing);
            document = null;
            //Close the file out
            fileOpen.Quit();
        }
        private void FindAndReplace(Application fileOpen, object findText, object replaceWithText)
        {
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = WdReplace.wdReplaceAll;
            object wrap = 1;
            //execute find and replace
            fileOpen.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }
        private void SearchReplace(Application fileOpen, object findText, object replaceWithText)
        {
            Find findObject = fileOpen.Selection.Find;
            findObject.ClearFormatting();
            findObject.Text = $"{findText}";
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = $"{replaceWithText}";
            object oMissing = System.Reflection.Missing.Value;
            object replaceAll = WdReplace.wdReplaceAll;
            findObject.Execute(oMissing,oMissing,oMissing,oMissing,oMissing,oMissing,oMissing,Replace: ref replaceAll);
        }
        private void SearchTextBox(Application fileOpen,string name,string newContent)
        {
            var a = fileOpen.Documents;
            foreach (Document b in a)
            {
                foreach (Shape shape in b.Shapes)
                    if (shape.Name.Contains("Text Box"))
                    {
                        if (shape.TextFrame.ContainingRange.Text.Contains(name))
                        {
                            shape.TextFrame.ContainingRange.Text = newContent;
                            return;
                        }
                        
                    }
            }
        }
    }
}