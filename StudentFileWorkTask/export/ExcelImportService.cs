using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using StudentFileWorkTask.presentation;

namespace StudentFileWorkTask.export
{
    internal class ExcelImportService
    {
        public FileAddResult ProcessAddFilesFromFolder(string folderPath, List<string> existingFiles)
        {
            string[] extensions = { "*.xlsx", "*.xls", "*.csv" };
            var files = new List<string>();
            foreach (var ext in extensions)
            {
                files.AddRange(System.IO.Directory.GetFiles(folderPath, ext));
            }

            var newFiles = new List<string>();
            int added = 0;
            foreach (var file in files)
            {
                if (!existingFiles.Contains(file))
                {
                    newFiles.Add(file);
                    added++;
                }
            }

            return new FileAddResult { NewFiles = newFiles, AddedCount = added };
        }
    }
}
