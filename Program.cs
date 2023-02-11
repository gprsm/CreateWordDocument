﻿using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
            Console.WriteLine("Enter Excel File path?");
            var excelPath = Console.ReadLine();
            while (string.IsNullOrEmpty(excelPath))
            {
                Console.WriteLine("inter Excel File path?");
                excelPath = Console.ReadLine();
            }

            var iDictionary = new Dictionary<int, PersonTypeNum.ColumnType>();
            Console.WriteLine("Enter name column Number...(begin from 0)");
            var nameColumnNumber = Console.ReadLine();
            while (string.IsNullOrEmpty(nameColumnNumber))
            {
                Console.WriteLine("Enter name column Number .....");
                nameColumnNumber = Console.ReadLine();
            }
            Console.WriteLine("Enter Family column Number .....");
            var familyColumnNumber = Console.ReadLine();
            while (string.IsNullOrEmpty(familyColumnNumber))
            {
                Console.WriteLine("Enter Family column Number .....");
                familyColumnNumber = Console.ReadLine();
            }
            iDictionary.Add(int.Parse(nameColumnNumber),PersonTypeNum.ColumnType.Name);
            iDictionary.Add(int.Parse(familyColumnNumber),PersonTypeNum.ColumnType.Family);
            
            IroxExcel iroxExcel = new IroxExcel();
            iroxExcel.ReadStyleSheet(iDictionary,excelPath);
            
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
