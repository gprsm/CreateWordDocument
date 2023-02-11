namespace CreateWordDocument.Models
{
    public class ExcelModel
    {
        public string NameAndFamily { get; set; }
        public string Company { get; set; }
        public PersonTypeNum.PersonType PersonType { get; set; }
        public PersonTypeNum.Gender Gender { get; set; }
        public string Text { get; set; }
        public string Signature { get; set; }
        public string Score { get; set; }
    }
}