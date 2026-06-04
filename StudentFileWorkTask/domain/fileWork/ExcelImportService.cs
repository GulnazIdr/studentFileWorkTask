using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using OfficeOpenXml;
using StudentFileWorkTask.presentation;

namespace TestReporter.domain.fileWork
{
    internal class ExcelImportService
    {

        public List<string> GetHeadersFromFile(string filePath)
        {
            var headers = new List<string>();
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.Columns;
                for (int col = 1; col <= colCount; col++)
                {
                    var header = worksheet.Cells[1, col].Text;
                    headers.Add(string.IsNullOrEmpty(header) ? $"Column{col}" : header);
                }
            }
            return headers;
        }
        public FileAddResult ProcessAddFilesFromFolder(string folderPath, List<string> existingFiles)
        {
            string[] extensions = { "*.xlsx", "*.xls", "*.csv" };
            var files = new List<string>();
            foreach (var ext in extensions)
            {
                files.AddRange(Directory.GetFiles(folderPath, ext));
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

        public List<string> LoadFilesFromFolder(string folderPath)
        {
            string[] extensions = { "*.xlsx", "*.xls", "*.csv" };
            var files = new List<string>();
            foreach (var ext in extensions)
            {
                files.AddRange(Directory.GetFiles(folderPath, ext));
            }

            if (files.Count == 0)
            {
                MessageBox.Show("В выбранной папке нет файлов Excel или CSV.");
                return files;
            }
            return files;
        }
    }
}
