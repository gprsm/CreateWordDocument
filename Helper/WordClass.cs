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
            string textPar,string folderPath)
        {
            var edited = false;
            if (!File.Exists(templatePath))
            {
                var info = new CollectInfo();
                templatePath=info.InteractWithUser("template not Found. enter full word file path again(example 'c:/template.docx')...");
            }
            object oMissing = System.Reflection.Missing.Value;
            Application fileOpen = new Application();
            Document document = fileOpen.Documents.Open(templatePath,ReadOnly:false);
            CreateDirectory createDirectory = new CreateDirectory();
            createDirectory.CreateSubDirectory(folderPath, "result");
            //Make the file visible 
            fileOpen.Visible = false;
            foreach (var excelInput in excelInputs)
            {
                //برای هر فرد
                var name = "";
                var family = "";
                var editedDoc = new Document();
                foreach (var model in excelInput.Models)
                {
                    var body = textPar;
                    if (!string.IsNullOrEmpty(model.Value))
                    {
                        switch (model.Type)
                        {
                            case PersonTypeNum.ColumnType.Name:
                            {
                                name = model.Value;
                                break;
                            }
                            case PersonTypeNum.ColumnType.Family:
                            {
                                family = model.Value;
                                break;
                            }
                            case PersonTypeNum.ColumnType.Score:
                            {
                                if (!string.IsNullOrEmpty(textWins))
                                {
                                    var place = textWins.IndexOf(model.PositionString, StringComparison.Ordinal);
                                    if (place>0)
                                    {
                                        var result = textWins.Remove(place, model.PositionString.Length).Insert(place, model.Value);
                                        body = result;
                                    }
                                }
                                break;
                            }
                            case PersonTypeNum.ColumnType.Gender:
                            {
                                if (model.Value== "مرد"|| model.Value=="مذکر"||
                                    model.Value=="1" ||
                                    model.Value=="پسر")
                                {
                                    editedDoc=SearchTextBox(document, model.PositionString, "جناب آقای");
                                    edited = true;
                                }

                                if (model.Value== "زن"|| model.Value=="مونث"||
                                    model.Value=="2" ||
                                    model.Value=="دختر")
                                {
                                    editedDoc=SearchTextBox(document, model.PositionString, "سرکار خانم");
                                    edited = true;
                                }

                                if (model.Value.Contains("الاسلام") || model.Value=="3")
                                {
                                    editedDoc=SearchTextBox(document, model.PositionString, "حجت الاسلام و المسلمین");
                                    edited = true;
                                }
                                break;
                            }
                            case PersonTypeNum.ColumnType.PersonType:
                            {
                                var textString = "";
                                if (model.Value==PersonTypeNum.PersonType.Colleague.ToString() ||
                                    model.Value.Contains("اصلی")|| model.Value.Contains("همکار"))
                                {
                                    textString = "همکار محترم";
                                }
                                editedDoc=SearchTextBox(document, model.PositionString,
                                    textString);

                                edited = true;
                                break;
                            }
                            case PersonTypeNum.ColumnType.Text:
                            {
                                editedDoc=SearchTextBox(document, model.PositionString,
                                    body);

                                edited = true;
                                break;
                            }
                        }
                    }

                    if (!edited)
                    {
                        editedDoc=SearchTextBox(document, model.PositionString, model.Value);
                    }
                    edited = false;
                }

                
                var resPath = $@"{folderPath}\result\{name}_{family}.docx";
                try
                {
                    editedDoc.SaveAs2(resPath);
                    object outputFileName = resPath.Replace(".docx", ".pdf");
                    var pdfPath =outputFileName;
                    object fileFormat = WdSaveFormat.wdFormatPDF;
                    document.SaveAs( pdfPath,
                        ref fileFormat, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
                
            }
            object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
            ((_Document)document).Close(ref saveChanges, ref oMissing, ref oMissing);
            document = null;
            //Close the file out
            fileOpen.Quit();
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
            SearchTextBox(document, textToR, replace);
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
        private Document SearchTextBox(Document doc,string name,string newContent)
        {
            foreach (Shape shape in doc.Shapes)
                if (shape.Name.Contains("Text Box"))
                {
                    if (shape.TextFrame.ContainingRange.Text.Contains(name))
                    {
                        shape.TextFrame.ContainingRange.Text = newContent;
                    }
                        
                }

            return doc;
        }
    }
}