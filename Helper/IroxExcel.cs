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
            foreach (var info in columnsInfo)
            {
                foreach (var totalRow in workSheet.Rows)
                {
                    var excelModel = new ExcelModel()
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
                    foreach (var cell in 
                             totalRow.ToList().Where(cell => 
                                 !string.IsNullOrEmpty(cell.Text)).
                                 Where(cell => cell.ColumnIndex==info.Key))
                    {
                        switch (info.Value.ColumnType)
                        {
                            case PersonTypeNum.ColumnType.Name:
                            {
                                excelModel.Name.PositionString = info.Value.PositionString;
                                excelModel.Name.Value = cell.Value.ToString();
                                break;
                            }
                            case PersonTypeNum.ColumnType.Family:
                            {
                                excelModel.Family.PositionString = info.Value.PositionString;
                                excelModel.Family.Value = cell.Value.ToString();
                                break;
                            }
                            case PersonTypeNum.ColumnType.PersonType:
                            {
                                var personType=(int)cell.Value==(int)PersonTypeNum.PersonType.Family?PersonTypeNum.PersonType.Family:
                                        PersonTypeNum.PersonType.Colleague;
                                excelModel.PersonType.PositionString = info.Value.PositionString;
                                excelModel.PersonType.Value = personType;
                                break;
                            }
                            case PersonTypeNum.ColumnType.Gender:
                            {
                                excelModel.Gender.PositionString = info.Value.PositionString;
                                if ((int)cell.Value==(int)PersonTypeNum.Gender.Man)
                                {
                                    excelModel.Gender.Value = PersonTypeNum.Gender.Man;
                                }
                                else
                                {
                                    var gender = (int)cell.Value==(int)PersonTypeNum.Gender.Woman ?
                                        PersonTypeNum.Gender.Woman : PersonTypeNum.Gender.Religious;
                                    excelModel.Gender.Value = gender;
                                }
                                break;
                            }
                            case PersonTypeNum.ColumnType.Text:
                            {
                                excelModel.Text.PositionString = info.Value.PositionString;
                                excelModel.Text.Value = cell.Value.ToString();
                                break;
                            }
                            case PersonTypeNum.ColumnType.Signature:
                            {
                                excelModel.Signature.PositionString = info.Value.PositionString;
                                excelModel.Signature.Value = cell.Value.ToString();
                                break;
                            }
                            case PersonTypeNum.ColumnType.Company:
                            {
                                excelModel.Company.PositionString = info.Value.PositionString;
                                excelModel.Company.Value = cell.Value.ToString();
                                break;
                            }
                            case PersonTypeNum.ColumnType.Score:
                            {
                                excelModel.Score.PositionString = info.Value.PositionString;
                                excelModel.Score.Value = cell.Value.ToString();
                                break;
                            }
                        }
                    }
                    resultList.Add(excelModel);
                }
            }
            return resultList;
        }
    }
}