using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CreateWordDocument.Models;
using Microsoft.Office.Interop.Word;

namespace CreateWordDocument.Helper
{
    public class WordClass
    {
        public void StartProcess(List<ExcelModel> excelInputs, string templatePath, string textWins,
            string textPar, string folderPath)
        {
            var counter = excelInputs.Count;
            foreach (var input in excelInputs)
            {
                ProcessInputs(input,templatePath,textWins,textPar,folderPath);
                --counter;
                Console.Write(counter);
            }
        }
        private void ProcessInputs(ExcelModel excelInput, string templatePath, string textWins,
            string textPar,string folderPath)
        {
            var edited = false;
            if (!File.Exists(templatePath))
            {
                var info = new CollectInfo();
                templatePath=info.InteractWithUser("template not Found. enter full word file path again(example 'c:/template.docx')...");
            }
            object oMissing = System.Reflection.Missing.Value;
            var fileOpen = new Application();
            var document = fileOpen.Documents.Open(templatePath,ReadOnly:false);
            var createDirectory = new CreateDirectory();
            createDirectory.CreateSubDirectory(folderPath, "result");
            //Make the file visible 
            fileOpen.Visible = false;
            
            //   بخش نام و نام خانوادگی و عنوان
            var name =excelInput.Models.FirstOrDefault(x => x.Type ==
                                                            PersonTypeNum.ColumnType.Name)?.Value;
            var family =excelInput.Models.FirstOrDefault(x => x.Type ==
                                                              PersonTypeNum.ColumnType.Family)?.Value;
            var genderValue=excelInput.Models.FirstOrDefault(x => x.Type ==
                                                                  PersonTypeNum.ColumnType.Gender)?.Value;
            var genderPositionString=excelInput.Models.FirstOrDefault(x => x.Type ==
                                                                           PersonTypeNum.ColumnType.Gender)?.PositionString;
            
            var personTitle = string.Empty;
            switch (genderValue)
            {
                case "مرد":
                case "مذکر":
                case "1":
                case "پسر":
                    personTitle = "جناب آقای";
                    break;
                case "زن":
                case "مونث":
                case "2":
                case "دختر":
                    personTitle = "سرکار خانم";
                    break;
            }

            if (genderValue!=null && genderValue.Contains("الاسلام") || genderValue=="3")
            {
                personTitle = "حجت الاسلام و مسلمین";
            }

            SearchTextBox(document, genderPositionString,$"{personTitle} {name} {family}");


            //بخش نوشتار اصلی و امتیازات
            var textPosition=excelInput.Models.FirstOrDefault(x => x.Type ==
                                                                   PersonTypeNum.ColumnType.Text)?.PositionString;
            var scoreValue=excelInput.Models.FirstOrDefault(x => x.Type ==
                                                            PersonTypeNum.ColumnType.Score)?.Value;
            var scorePositionString=excelInput.Models.FirstOrDefault(x => x.Type ==
                                                                 PersonTypeNum.ColumnType.Score)?.PositionString;
            if (!string.IsNullOrEmpty(scoreValue) && !string.IsNullOrEmpty(textWins) && scorePositionString != null)
            {
                var place = textWins.IndexOf(scorePositionString, StringComparison.Ordinal);
                var textBody = place>0 ? textWins.Remove(place, scorePositionString.Length).Insert(place, scoreValue) : textPar;
                SearchTextBox(document, textPosition, textBody);
            }
            //بخش عنوان کاری
            
            var personType=excelInput.Models.FirstOrDefault(x => x.Type ==
                                                                 PersonTypeNum.ColumnType.PersonType)?.Value;
            var personTypePositionString=excelInput.Models.FirstOrDefault(x => x.Type ==
                                                                               PersonTypeNum.ColumnType.PersonType)?.PositionString;
            var personTypeString=String.Empty;
            if (personType=="2" ||
                personType.Contains("اصلی")|| personType.Contains("همکار"))
            {
                personTypeString = "همکار محترم";
            }
            var companyStr = excelInput.Models.FirstOrDefault(x => x.Type ==
                                                                   PersonTypeNum.ColumnType.Company)?.Value;
            if (!string.IsNullOrEmpty(personTypeString))
            {
                SearchTextBox(document, personTypePositionString, $"{personTypeString} {companyStr}");
            }
            
            var resPath = $@"{folderPath}\result\{name} {family}.docx";
                try
                {
                    document.SaveAs2(resPath);
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