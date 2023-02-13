using System.Collections.Generic;

namespace CreateWordDocument.Models
{
    public class ExcelModel
    {
        public IDictionary<string,string> Name { get; set; }
        public IDictionary<string,string> Family { get; set; }
        public IDictionary<string,string> Company { get; set; }
        public IDictionary<string,PersonTypeNum.PersonType> PersonType { get; set; }
        public IDictionary<string,PersonTypeNum.Gender> Gender { get; set; }
        public IDictionary<string,string> Text { get; set; }
        public IDictionary<string,string> Signature { get; set; }
        public IDictionary<string,string> Score { get; set; }
    }
}