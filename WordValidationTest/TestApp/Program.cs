using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordValidationTest;

namespace TestApp
{
    class Program
    {
        static void Main(string[] args)
        {
            WordValidator validator = new WordValidator("Test Document.docx", FileFormatVersions.Office2016);
            List<ValidationErrorInfo> errorInfo = validator.ValidateDocument();
            validator.Dispose();
            if (errorInfo == null || errorInfo.Count > 0)
            {
                Console.WriteLine("Validation Errors!");

                foreach (var error in errorInfo)
                    Console.WriteLine(error.Description);
            }
            else
            {
                Console.WriteLine("No Validation Errors!");
            }

            Console.ReadKey();
        }
    }
}
