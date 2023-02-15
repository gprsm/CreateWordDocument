
using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;

namespace CreateWordDocument.Models
{
    public class ExcelModel
    {
        public List<DefinedValue<string>> Models{ get; set; }
    }

    public class DefinedValue <T>
    {
        public string PositionString{ get; set; }
        public T Value { get; set; }
        public PersonTypeNum.ColumnType Type{ get; set; }
    }
}