using System.IO;

namespace CreateWordDocument.Helper
{
    public class ReadTextFile
    {
        public string ReadText(string textPath)
        {
            var result = "";
            if (!File.Exists(textPath))
            {
                var info = new CollectInfo();
                textPath=info.InteractWithUser("File not Found. enter full text file path again(example 'c:/sample.txt')...");
            }
            // Store each line in array of strings
            string[] lines = File.ReadAllLines(textPath);

            foreach (var ln in lines)
            {
                if (string.IsNullOrEmpty(result))
                {
                    result = ln;
                }
                else
                {
                    result += "\n" + ln;
                }
            }
            return result;
        }
    }
}