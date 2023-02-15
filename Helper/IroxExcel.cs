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
                    Models = new List<DefinedValue<string>>()
                };
                foreach (var model in columnsInfo)
                {
                    var cellInfo = sheetRow.FirstOrDefault(x => x.ColumnIndex == model.Key);
                    if (cellInfo?.Value!=null)
                    {
                        excel.Models.Add(new DefinedValue<string>()
                        {
                            PositionString = model.Value.PositionString,
                            Value = cellInfo.Value.ToString(),
                            Type = model.Value.ColumnType
                        });
                    }
                }
                resultList.Add(excel);
            }
            return resultList;
        }
    }
}