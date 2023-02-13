using System;
using CreateWordDocument.Models;

namespace CreateWordDocument.Helper
{
    public class CollectInfo
    {
        public string InteractWithUser(string message)
        {
            Console.WriteLine($"{message}");
            var excelPath = $@"{Console.ReadLine()}";
            while (string.IsNullOrEmpty(excelPath))
            {
                Console.WriteLine($"{message}");
                excelPath = Console.ReadLine();
            }
            return excelPath;
        }

        public ProcessInfoModel CallUser()
        {
            return new ProcessInfoModel
            {
                FolderPath = InteractWithUser("Enter folder of FilesPath(excel and word template)..."),
                ExcelName = InteractWithUser("Enter Excel file name..."),
                WordTemplateName = InteractWithUser("Enter word template file name..."),
                NameColumnNumber = int.Parse(InteractWithUser("Enter number of Name Column...")),
                NameCharInTemplate = InteractWithUser("Enter Character of Name string..."),
                FamilyColumnNumber = int.Parse(InteractWithUser("Enter number of family Column...")),
                FamilyCharInTemplate = InteractWithUser("Enter Character of family string..."),
                PersonTypeColumnNumber = int.Parse(InteractWithUser("Enter number of Column that to define her/him is Family or Colleague...")),
                PersonTypeCharInTemplate = InteractWithUser("Enter Character of this PersonType(family/Colleague) string in template..."),
                GenderColumnNumber = int.Parse(InteractWithUser("Enter number of Gender Column...")),
                GenderCharInTemplate = InteractWithUser("Enter Character of Gender string..."),
                CompanyColumnNumber = int.Parse(InteractWithUser("Enter number of Company Column...")),
                CompanyCharInTemplate = InteractWithUser("Enter Character of Company string..."),
                ScoreColumnNumber = int.Parse(InteractWithUser("Enter number of Score Column...")),
                ScoreCharInTemplate = InteractWithUser("Enter Character of Score string..."),
                SignatureColumnNumber = int.Parse(InteractWithUser("Enter number of Signature Column...")),
                SignatureCharInTemplate = InteractWithUser("Enter Character of Signature string..."),
                TextWinsName = InteractWithUser("Enter Text boy for wins File Name..."),
                TextParticipantsName = InteractWithUser("Enter Text boy for Participants File Name..."),
                TextCharInTemplate = InteractWithUser("Enter Character of text string...")
            };
        }
    }
}