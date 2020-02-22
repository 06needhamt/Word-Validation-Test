using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordValidationTest
{
    /// <summary>
    /// Class Used to Verify Word Documents
    /// </summary>
    public class WordValidator : OpenXmlValidator, IDisposable
    {
        /// <summary>
        /// Path of the File Bieng Validated
        /// </summary>
        public string FilePath { get; set; }

        /// <summary>
        /// Is the File Valid True/False
        /// </summary>
        public bool Valiated { get; set; }

        /// <summary>
        /// Detailed Validation Error Implementation 
        /// </summary>
        public List<ValidationErrorInfo> ErrorInfo { get; private set; }

        /// <summary>
        /// Validation Error Types 
        /// </summary>
        public List<ValidationErrorType> ErrorType { get; private set; }

        public OpenXmlPackage Package { get; set; }

        public WordValidator(string filePath, FileFormatVersions version) : base((FileFormatVersions)version)
        {
            this.FilePath = filePath;
            this.MaxNumberOfErrors = 1;
            this.Valiated = false;
            this.Package = WordprocessingDocument.Open(filePath, false);
            //this.FileFormat = (DocumentFormat.OpenXml.FileFormatVersions)version;
        }

        public List<ValidationErrorInfo> ValidateDocument(OpenXmlPackage package)
        {
            List<ValidationErrorInfo> validationErrors = this.Validate(package).ToList();
            if (validationErrors == null || validationErrors.Count == 0)
            {
                this.Valiated = true;
                this.ErrorInfo = null;
                this.ErrorType = null;
            }
            else
            {
                this.Valiated = false;
                this.ErrorInfo = validationErrors;
                this.ErrorType = validationErrors.Select(x => x.ErrorType).Cast<ValidationErrorType>().ToList();
            }
            return validationErrors;
        }

        public List<ValidationErrorInfo> ValidateDocument()
        {
            List<ValidationErrorInfo> validationErrors = this.Validate(Package).ToList();
            if (validationErrors == null || validationErrors.Count == 0)
                this.Valiated = true;
            return validationErrors;
        }

        public void Dispose()
        {
            Package.Dispose();
        }
    }
}
