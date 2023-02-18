using System;
using System.Collections.Generic;
using CreateWordDocument.Helper;
using CreateWordDocument.Models;

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
                WordTemplateName = "doc1",
                NameColumnNumber = 1,
                NameCharInTemplate = "name",
                FamilyColumnNumber = 2,
                FamilyCharInTemplate = "family",
                PersonTypeColumnNumber = 3,
                PersonTypeCharInTemplate = "pt",
                GenderColumnNumber = 4,
                GenderCharInTemplate = "gen",
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
                },
                {
                    0, new PositionAndTypeModel()
                    {
                    ColumnType = PersonTypeNum.ColumnType.Text,
                    PositionString = processInfo.TextCharInTemplate
                    } 
                }
            };
            IronExcel ironExcel = new IronExcel();
            var result=ironExcel.ReadStyleSheet(iDictionary,$@"{processInfo.FolderPath}/{processInfo.ExcelName}.xlsx");
            ReadTextFile textFile = new ReadTextFile();
            var bodyWins = textFile.ReadText($@"{processInfo.FolderPath}/{processInfo.TextWinsName}.txt");
            var bodyPart = textFile.ReadText($@"{processInfo.FolderPath}/{processInfo.TextParticipantsName}.txt");
            WordClass wordClass = new WordClass();
            wordClass.StartProcess(result,$@"{processInfo.FolderPath}/{processInfo.WordTemplateName}.docx",bodyWins,bodyPart,$@"{processInfo.FolderPath}");

            Console.WriteLine("Finish...");
        }
    }
}
