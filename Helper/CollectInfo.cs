using System;
using System.Collections.Generic;
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
    }
}