using System;
using System.IO;
using System.Collections.Generic;
using CreateWordDocument.Helper;
using CreateWordDocument.Models;
using DocumentFormat.OpenXml.Packaging;
using InaOfficeTools;

namespace CreateWordDocument
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Init...");
            var info = new CollectInfo();
            //var processInfo = info.CallUser();
            var processInfo = new ProcessInfoModel
            {
                FolderPath = @"C:\Users\mohse\Desktop\New_folder",
                ExcelName = "Book1",
                WordTemplateName = "Template1",
                NameColumnNumber = 1,
                NameCharInTemplate = "name",
                FamilyColumnNumber = 2,
                FamilyCharInTemplate = "family",
                PersonTypeColumnNumber = 3,
                PersonTypeCharInTemplate = "perType",
                GenderColumnNumber = 4,
                GenderCharInTemplate = "gender",
                CompanyColumnNumber = 5,
                CompanyCharInTemplate = "comp",
                ScoreColumnNumber = 6,
                ScoreCharInTemplate = "score",
                SignatureColumnNumber = 7,
                SignatureCharInTemplate = "sign",
                TextWinsName = "wins",
                TextParticipantsName = "par",
                TextCharInTemplate = "text"
            };

             var iDictionary = new Dictionary<int, PositionAndTypeModel>
             {
                {
                    processInfo.NameColumnNumber, new PositionAndTypeModel()
                    {
                        ColumnType = PersonTypeNum.ColumnType.Name,
                        PositionString = processInfo.NameCharInTemplate
                    }
                },
                {
                    processInfo.FamilyColumnNumber, new PositionAndTypeModel()
                    {
                        ColumnType = PersonTypeNum.ColumnType.Family,
                        PositionString = processInfo.FamilyCharInTemplate
                    }
                },
                { 
                    processInfo.PersonTypeColumnNumber, new PositionAndTypeModel()
                    {
                    ColumnType = PersonTypeNum.ColumnType.PersonType,
                    PositionString = processInfo.PersonTypeCharInTemplate
                    }
                },
                { 
                    processInfo.GenderColumnNumber, new PositionAndTypeModel()
                    {
                    ColumnType = PersonTypeNum.ColumnType.Gender,
                    PositionString = processInfo.GenderCharInTemplate
                    }
                },
                { 
                    processInfo.CompanyColumnNumber, new PositionAndTypeModel()
                    {
                    ColumnType = PersonTypeNum.ColumnType.Company,
                    PositionString = processInfo.CompanyCharInTemplate
                    }
                },
                { 
                    processInfo.ScoreColumnNumber, new PositionAndTypeModel()
                    {
                    ColumnType = PersonTypeNum.ColumnType.Score,
                    PositionString = processInfo.ScoreCharInTemplate
                    }
                },
                { 
                    processInfo.SignatureColumnNumber, new PositionAndTypeModel()
                    {
                    ColumnType = PersonTypeNum.ColumnType.Signature,
                    PositionString = processInfo.SignatureCharInTemplate
                    } 
                }
            };
            IronExcel ironExcel = new IronExcel();
            var result=ironExcel.ReadStyleSheet(iDictionary,$@"{processInfo.FolderPath}/{processInfo.ExcelName}.xlsx");
            ReadTextFile textFile = new ReadTextFile();
            var bodyWins = textFile.ReadText($@"{processInfo.FolderPath}/{processInfo.TextWinsName}.txt");
            var bodyPart = textFile.ReadText($@"{processInfo.FolderPath}/{processInfo.TextParticipantsName}.txt");
            WordClass wordClass = new WordClass();
            string documentFolder= @"C:\Users\mohse\Desktop\New_folder\{0}";
            wordClass.FindAndReplace(string.Format(documentFolder,"Template1.docx"),"family","احمد نصیری",string.Format(documentFolder,"TemplateResult.docx"));
            FileStream fileStream = new FileStream(string.Format(documentFolder,"Template1.docx"), FileMode.Open);
           
          
            using (MemoryStream memStr = new MemoryStream())
            {
                fileStream.CopyTo(memStr);
                fileStream.Close();
                using (WordprocessingDocument WPDoc = WordprocessingDocument.Open(memStr, true))
                {
                    Console.WriteLine("Creating...");
                    WordGenerator objWord = new WordGenerator(WPDoc);
                    
                    //Inserting text 
                    objWord.UpdateTextoControlWord("PropName","Juan Alberto Zapata Suarez" );
                    objWord.UpdateTextoControlWord("PropAge", "35 años");
                    objWord.UpdateTextoControlWord("PropDate", DateTime.Now.ToString("dd/MMMM/yyy"));
                    

                    //inserting bullets
                    List<BulletsConfigWordGenerator> bulletsList = new List<BulletsConfigWordGenerator>();
                    bulletsList.Add(new BulletsConfigWordGenerator("Power platform", 0, true, true, "41"));
                    bulletsList.Add(new BulletsConfigWordGenerator("Power BI", 1, false, false, "21"));
                    bulletsList.Add(new BulletsConfigWordGenerator("Chartuculator", 2, false, false, "21"));
                    bulletsList.Add(new BulletsConfigWordGenerator("spfx-pbiviz ", 2, false, false, "21"));
                    bulletsList.Add(new BulletsConfigWordGenerator("Power Automate", 1, false, false, "21"));
                    bulletsList.Add(new BulletsConfigWordGenerator("Power Apps", 1, false, false, "21"));
                    bulletsList.Add(new BulletsConfigWordGenerator("Programing", 0, true, true, "41"));
                    bulletsList.Add(new BulletsConfigWordGenerator("C#", 1, false, false, "21"));
                    bulletsList.Add(new BulletsConfigWordGenerator("Java Script", 1, false, false, "21"));
                    objWord.UpdateBulletsControlWord("PropBullets", bulletsList);

                    //inserting table
                    TableConfigWordGenerator tableConf = new TableConfigWordGenerator() { RowList= new List<RowConfigWordGenerator>() };                   

                    RowConfigWordGenerator objRow1 = new RowConfigWordGenerator();
                    objRow1.CellList = new List<CellConfigWordGenerator>();
                    objRow1.CellList.Add(new CellConfigWordGenerator() { Text = "______________________________" });
                    objRow1.CellList.Add(new CellConfigWordGenerator() { Text = "______________________________" });
                    tableConf.RowList.Add(objRow1);

                    RowConfigWordGenerator objRow2 = new RowConfigWordGenerator();
                    objRow2.CellList = new List<CellConfigWordGenerator>();
                    objRow2.CellList.Add(new CellConfigWordGenerator() { Text = "CEO. JOSE MARTINEZ" });
                    objRow2.CellList.Add(new CellConfigWordGenerator() { Text = "CTO. CARLOS GONZALEZ" });
                    tableConf.RowList.Add(objRow2);

                    objWord.UpdateTablaControlWord("TBLSignature", tableConf);
                }

                FileStream file = new FileStream(string.Format(documentFolder, "FinalWord.docx"), FileMode.Create, FileAccess.Write);
                memStr.WriteTo(file);
                file.Close();
            }

            Console.WriteLine("Finish...");
        }
    }
}
