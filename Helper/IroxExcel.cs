using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CreateWordDocument.Models;
using IronXL;

namespace CreateWordDocument.Helper
{
    public class IronExcel
    {
        public List<ExcelModel> ReadStyleSheet(IDictionary<int,PositionAndTypeModel> columnsInfo,string xmlPath)
        {
            if (!File.Exists(xmlPath))
            {
                var info = new CollectInfo();
                xmlPath=info.InteractWithUser("File not Found. enter full excel file path again(example 'c:/sample.xlsx')...");
            }
            var resultList = new List<ExcelModel>();
            WorkBook workBook = WorkBook.Load(xmlPath);
            WorkSheet workSheet = workBook.WorkSheets.First();
            
            foreach (var sheetRow in workSheet.Rows)
            {
                var excel = new ExcelModel()
                {
                    Text = new DefinedValue<string>(),
                    Company = new DefinedValue<string>(),
                    Family = new DefinedValue<string>(),    
                    Gender = new DefinedValue<PersonTypeNum.Gender>(),
                    Name = new DefinedValue<string>(),
                    Score = new DefinedValue<string>(),
                    PersonType = new DefinedValue<PersonTypeNum.PersonType>(),
                    Signature = new DefinedValue<string>()
                };
                foreach (var model in columnsInfo)
                {
                    var cellInfo = sheetRow.FirstOrDefault(x => x.ColumnIndex == model.Key);
                    if (cellInfo?.Value!=null)
                    {
                        switch (model.Value.ColumnType)
                        {
                            case PersonTypeNum.ColumnType.Name:
                            {
                                excel.Name.PositionString = model.Value.PositionString;
                                excel.Name.Value = cellInfo.Value.ToString();
                                break;
                            }
                            case PersonTypeNum.ColumnType.Family:
                            {
                                excel.Family.PositionString = model.Value.PositionString;
                                excel.Family.Value = cellInfo.Value.ToString();
                                break;
                            }
                            case PersonTypeNum.ColumnType.PersonType:
                            {
                                try
                                {
                                    var personType=int.Parse(cellInfo.Value.ToString())==(int)PersonTypeNum.PersonType.Family?PersonTypeNum.PersonType.Family:
                                        PersonTypeNum.PersonType.Colleague;
                                    excel.PersonType.PositionString = model.Value.PositionString;
                                    excel.PersonType.Value = personType;
                                }
                                catch (Exception)
                                {
                                    excel.PersonType.PositionString = model.Value.PositionString;
                                    excel.PersonType.Value = PersonTypeNum.PersonType.Colleague;
                                }
                                break;
                            }
                            case PersonTypeNum.ColumnType.Gender:
                            {
                                excel.Gender.PositionString = model.Value.PositionString;
                                try
                                {
                                    if (int.Parse(cellInfo.Value.ToString())==(int)PersonTypeNum.Gender.Man)
                                    {
                                        excel.Gender.Value = PersonTypeNum.Gender.Man;
                                    }
                                    else
                                    {
                                        var gender = int.Parse(cellInfo.Value.ToString())==(int)PersonTypeNum.Gender.Woman ?
                                            PersonTypeNum.Gender.Woman : PersonTypeNum.Gender.Religious;
                                        excel.Gender.Value = gender;
                                    }
                                }
                                catch (Exception)
                                {
                                    excel.Gender.Value = PersonTypeNum.Gender.Man;
                                }
                                break;
                            }
                            case PersonTypeNum.ColumnType.Text:
                            {
                                excel.Text.PositionString = model.Value.PositionString;
                                excel.Text.Value = cellInfo.Value.ToString();
                                break;
                            }
                            case PersonTypeNum.ColumnType.Signature:
                            {
                                excel.Signature.PositionString = model.Value.PositionString;
                                excel.Signature.Value = cellInfo.Value.ToString();
                                break;
                            }
                            case PersonTypeNum.ColumnType.Company:
                            {
                                excel.Company.PositionString = model.Value.PositionString;
                                excel.Company.Value = cellInfo.Value.ToString();
                                break;
                            }
                            case PersonTypeNum.ColumnType.Score:
                            {
                                excel.Score.PositionString = model.Value.PositionString;
                                excel.Score.Value = cellInfo.Value.ToString();
                                break;
                            }
                        }
                    }
                }
                resultList.Add(excel);
            }
            return resultList;
        }
    }
}