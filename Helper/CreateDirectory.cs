using System;
using System.IO;

namespace CreateWordDocument.Helper
{
    public class CreateDirectory
    {
        public bool CreateSubDirectory(string oldPath,string newSubFolder)
        {
            try
            {
                if (!Directory.Exists($@"{oldPath}\{newSubFolder}"))
                {
                    Directory.CreateDirectory($@"{oldPath}\{newSubFolder}");
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}