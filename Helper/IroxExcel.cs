using System.Collections.Generic;
using System.Linq;
using CreateWordDocument.Models;
using IronXL;

namespace CreateWordDocument.Helper
{
    public class IronExcel
    {
        public List<ExcelModel> ReadStyleSheet(IDictionary<int,PositionAndTypeModel> columnsInfo,string xmlPath)
        {
            var resultList = new List<ExcelModel>();
            WorkBook workBook = WorkBook.Load(xmlPath);
            WorkSheet workSheet = workBook.WorkSheets.First();
            foreach (var info in columnsInfo)
            {
                foreach (var totalRow in workSheet.Rows)
                {
                    var excelModel = new ExcelModel();
                    foreach (var cell in 
                             totalRow.ToList().Where(cell => 
                                 !string.IsNullOrEmpty(cell.Text)).
                                 Where(cell => cell.ColumnIndex==info.Key))
                    {
                        switch (info.Value.ColumnType)
                        {
                            case PersonTypeNum.ColumnType.Name:
                            {
                                excelModel.Name.Add(info.Value.PositionString,cell.Value.ToString()); 
                                break;
                            }
                            case PersonTypeNum.ColumnType.Family:
                            {
                                
                                break;
                            }
                            case PersonTypeNum.ColumnType.PersonType:
                            {
                                var personTypeNum=(int)cell.Value==(int)PersonTypeNum.PersonType.Family?PersonTypeNum.PersonType.Family:
                                        PersonTypeNum.PersonType.Colleague;
                                excelModel.PersonType.Add(info.Value.PositionString,personTypeNum); 
                                break;
                            }
                            case PersonTypeNum.ColumnType.Gender:
                            {
                                if ((int)cell.Value==(int)PersonTypeNum.Gender.Man)
                                {
                                    excelModel.Gender.Add(info.Value.PositionString,PersonTypeNum.Gender.Man);
                                }
                                else
                                {
                                    var gender = (int)cell.Value==(int)PersonTypeNum.Gender.Woman ?
                                        PersonTypeNum.Gender.Woman : PersonTypeNum.Gender.Religious;
                                    excelModel.Gender.Add(info.Value.PositionString,gender);
                                }
                                break;
                            }
                            case PersonTypeNum.ColumnType.Text:
                            {
                                excelModel.Text.Add(info.Value.PositionString,(string)cell.Value);
                                break;
                            }
                            case PersonTypeNum.ColumnType.Signature:
                            {
                                excelModel.Signature.Add(info.Value.PositionString,(string)cell.Value);
                                break;
                            }
                            case PersonTypeNum.ColumnType.Company:
                            {
                                excelModel.Company.Add(info.Value.PositionString,(string)cell.Value);
                                break;
                            }
                            case PersonTypeNum.ColumnType.Score:
                            {
                                excelModel.Score.Add(info.Value.PositionString,(string)cell.Value);
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