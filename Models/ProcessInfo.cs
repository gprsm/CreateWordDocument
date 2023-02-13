namespace CreateWordDocument.Models
{
    public class ProcessInfoModel
    {
        public string FolderPath { get; set; }
        public string ExcelName { get; set; }
        public string WordTemplateName { get; set; }
        public string TextCharInTemplate { get; set; }
        public string TextParticipantsName { get; set; }
        public string TextWinsName { get; set; }
        public string SignatureCharInTemplate { get; set; }
        public int SignatureColumnNumber { get; set; }  
        public string ScoreCharInTemplate { get; set; }
        public int ScoreColumnNumber { get; set; }
        public string CompanyCharInTemplate { get; set; }
        public int CompanyColumnNumber { get; set; }
        public string GenderCharInTemplate { get; set; }
        public int GenderColumnNumber { get; set; }
        public string PersonTypeCharInTemplate { get; set; }
        public int PersonTypeColumnNumber { get; set; }
        public string FamilyCharInTemplate { get; set; }
        public int FamilyColumnNumber { get; set; }
        public string NameCharInTemplate { get; set; }
        public int NameColumnNumber { get; set; }
    }
}