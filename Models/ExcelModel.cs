
namespace CreateWordDocument.Models
{
    public class ExcelModel
    {
        public DefinedValue<string> Name { get; set; }
        public DefinedValue<string> Family { get; set; }
        public DefinedValue<string> Company { get; set; }
        public DefinedValue<PersonTypeNum.PersonType> PersonType { get; set; }
        public DefinedValue<PersonTypeNum.Gender> Gender { get; set; }
        public DefinedValue<string> Text { get; set; }
        public DefinedValue<string> Signature { get; set; }
        public DefinedValue<string> Score { get; set; }
    }

    public class DefinedValue <T>
    {
        public string PositionString{ get; set; }
        public T Value { get; set; }
    }
}