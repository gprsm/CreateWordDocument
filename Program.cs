using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
            CollectInfo info = new CollectInfo();
            var excelPath=info.InteractWithUser("Enter Excel FilePath...");
            var nameColumnNumber=int.Parse(info.InteractWithUser("Enter number of Name Column..."));
            var nameCharInTemplate=info.InteractWithUser("Enter Character of Name string...");
            var familyColumnNumber=int.Parse(info.InteractWithUser("Enter number of family Column..."));
            var familyCharInTemplate=info.InteractWithUser("Enter Character of family string...");
            var personTypeColumnNumber=int.Parse(info.InteractWithUser("Enter number of Column that to define her/him is Family or Colleague..."));
            var personTypeCharInTemplate=info.InteractWithUser("Enter Character of this PersonType(family/Colleague) string in template...");
            var genderColumnNumber=int.Parse(info.InteractWithUser("Enter number of Gender Column..."));
            var genderCharInTemplate=info.InteractWithUser("Enter Character of Gender string...");
            var companyColumnNumber=int.Parse(info.InteractWithUser("Enter number of Company Column..."));
            var companyCharInTemplate=info.InteractWithUser("Enter Character of Company string...");
            var scoreColumnNumber=int.Parse(info.InteractWithUser("Enter number of Score Column..."));
            var scoreCharInTemplate=info.InteractWithUser("Enter Character of Score string...");
            var signatureColumnNumber=int.Parse(info.InteractWithUser("Enter number of Signature Column..."));
            var signatureCharInTemplate=info.InteractWithUser("Enter Character of Signature string...");
            var textWinsPath=info.InteractWithUser("Enter Text boy for wins FilePath...");
            var textParticipantsPath=info.InteractWithUser("Enter Text boy for Participants FilePath...");
            var textCharInTemplate=info.InteractWithUser("Enter Character of text string...");
            
            var iDictionary = new Dictionary<int, PositionAndTypeModel>
            {
                {
                    nameColumnNumber, new PositionAndTypeModel()
                    {
                        ColumnType = PersonTypeNum.ColumnType.Name,
                        PositionString = nameCharInTemplate
                    }
                },
                {
                    familyColumnNumber, new PositionAndTypeModel()
                    {
                        ColumnType = PersonTypeNum.ColumnType.Family,
                        PositionString = familyCharInTemplate
                    }
                },
                { personTypeColumnNumber, new PositionAndTypeModel()
                {
                    ColumnType = PersonTypeNum.ColumnType.PersonType,
                    PositionString = personTypeCharInTemplate
                } },
                { genderColumnNumber, new PositionAndTypeModel()
                {
                    ColumnType = PersonTypeNum.ColumnType.Gender,
                    PositionString = genderCharInTemplate
                } },
                { companyColumnNumber, new PositionAndTypeModel()
                {
                    ColumnType = PersonTypeNum.ColumnType.Company,
                    PositionString = companyCharInTemplate
                } },
                { scoreColumnNumber, new PositionAndTypeModel()
                {
                    ColumnType = PersonTypeNum.ColumnType.Score,
                    PositionString = scoreCharInTemplate
                } },
                { signatureColumnNumber, new PositionAndTypeModel()
                {
                    ColumnType = PersonTypeNum.ColumnType.Signature,
                    PositionString = signatureCharInTemplate
                } }
            };
            IronExcel ironExcel = new IronExcel();
            var result=ironExcel.ReadStyleSheet(iDictionary,excelPath);
            ReadTextFile textFile = new ReadTextFile();
            var bodyWins = textFile.ReadText($@"{textWinsPath}");
            var bodyPart = textFile.ReadText($@"{textParticipantsPath}");
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
