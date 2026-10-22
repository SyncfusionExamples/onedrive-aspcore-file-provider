using System.Collections.Generic;

namespace Syncfusion.Web.FileManager.Base
{
    public class ErrorDetails
    {
        public string Code { get; set; }

        public string Message { get; set; }

        public IEnumerable<string> FileExists { get; set; }
    }
}
