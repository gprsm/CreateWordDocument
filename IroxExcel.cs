using System.Collections.Generic;
using System.Linq;
using CreateWordDocument.Models;
using IronXL;
namespace CreateWordDocument
{
    public class IroxExcel
    {
        public List<ExcelModel> ReadStyleSheet(IDictionary<int,PersonTypeNum.ColumnType> columnsNumber,string xmlPath)
        {
            var resultList = new List<ExcelModel>();
            WorkBook workBook = WorkBook.Load(xmlPath);
            WorkSheet workSheet = workBook.WorkSheets.First();
            foreach (var number in columnsNumber)
            {
                foreach (var totalRow in workSheet.Rows)
                {
                    var excelModel = new ExcelModel();
                    foreach (var cell in 
                             totalRow.ToList().Where(cell => 
                                 !string.IsNullOrEmpty(cell.Text)).
                                 Where(cell => cell.ColumnIndex==number.Key))
                    {
                        switch (number.Value)
                        {
                            case PersonTypeNum.ColumnType.Name:
                            {
                                excelModel.NameAndFamily = cell.Value.ToString();
                                break;
                            }
                            case PersonTypeNum.ColumnType.Family:
                            {
                                excelModel.NameAndFamily += " "+cell.Value.ToString();
                                break;
                            }
                            case PersonTypeNum.ColumnType.PersonType:
                            {
                                excelModel.PersonType = 
                                    (string)cell.Value=="1"?PersonTypeNum.PersonType.Colleague:
                                        PersonTypeNum.PersonType.Family;
                                break;
                            }
                            case PersonTypeNum.ColumnType.Gender:
                            {
                                if ((int)cell.Value==(int)PersonTypeNum.Gender.Man)
                                {
                                    excelModel.Gender = PersonTypeNum.Gender.Man;
                                }
                                else
                                {
                                    excelModel.Gender = (int)cell.Value==(int)PersonTypeNum.Gender.Woman ?
                                        PersonTypeNum.Gender.Woman : PersonTypeNum.Gender.Religious;
                                }
                                break;
                            }
                            case PersonTypeNum.ColumnType.Text:
                            {
                                excelModel.Text = (string)cell.Value;
                                break;
                            }
                            case PersonTypeNum.ColumnType.Signature:
                            {
                                excelModel.Signature = (string)cell.Value;
                                break;
                            }
                            case PersonTypeNum.ColumnType.Company:
                            {
                                excelModel.Company = (string)cell.Value;
                                break;
                            }
                            case PersonTypeNum.ColumnType.Score:
                            {
                                excelModel.Score = (string)cell.Value;
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