using System.Collections.Generic;

namespace Syncfusion.Web.FileManager.Base
{
    /// <summary>
    /// Represents the response returned by the file manager service.
    /// </summary>
    public class FileManagerResponse
    {
        /// <summary>
        /// Gets or sets the current working directory content.
        /// </summary>
        public FileManagerDirectoryContent CWD { get; set; }

        /// <summary>
        /// Gets or sets the collection of files/folders in the current directory.
        /// </summary>
        public IEnumerable<FileManagerDirectoryContent> Files { get; set; }

        /// <summary>
        /// Gets or sets the details of any errors encountered during the operation.
        /// </summary>
        public ErrorDetails Error { get; set; }

        /// <summary>
        /// Gets or sets additional details about the file or directory.
        /// </summary>
        public FileDetails Details { get; set; }
    }
}
