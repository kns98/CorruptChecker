using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml.Core.ExcelPackage;
using PdfSharp.Pdf.IO;
using SixLabors.ImageSharp;

namespace FileChecker
{
    class Program
    {
        static void Main(string[] args)
        {
            string basePath = args.Length > 0 ? args[0] : "./";
            List<string> corruptedFiles = SearchFiles(basePath);
            if (corruptedFiles.Count > 0)
            {
                ConfirmAndDeleteFiles(corruptedFiles);
            }
            else
            {
                Console.WriteLine("No corrupted files found.");
            }
        }

        static List<string> SearchFiles(string basePath)
        {
            Console.WriteLine($"Checking files in directory: {basePath}");
            var corruptedFiles = new List<string>();

            var types = new List<(string[] exts, string fileType)>
            {
                (new[] { ".png", ".jpg", "jpeg" }, "image"),
                (new[] { ".pdf" }, "pdf"),
                (new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".odf", ".ods", ".odt" }, "excel")
            };

            foreach (var filePath in Directory.EnumerateFiles(basePath))
            {
                Console.WriteLine(new string('-', 30));
                Console.WriteLine($"Checking file: {filePath}");

                bool? check = null;

                foreach (var option in types)
                {
                    if (option.exts.Any(ext => filePath.EndsWith(ext, StringComparison.OrdinalIgnoreCase)))
                    {
                        check = CheckFile(filePath, option.fileType);
                        break;
                    }
                }

                if (check == null)
                {
                    check = CheckFile(filePath, null);
                }

                if (check == false)
                {
                    Console.WriteLine("File is corrupted!");
                    corruptedFiles.Add(filePath);
                }
                else if (check == true)
                {
                    Console.WriteLine("File is safe!");
                }
                else
                {
                    Console.WriteLine("Could not verify the file.");
                }
            }

            return corruptedFiles;
        }

        static bool? CheckFile(string filePath, string fileType)
        {
            bool? check = null;

            try
            {
                if (fileType == "image")
                {
                    using var img = Image.Load(filePath);
                    check = true;
                }
                else if (fileType == "pdf")
                {
                    using var pdf = PdfReader.Open(filePath, PdfDocumentOpenMode.InformationOnly);
                    check = pdf.Info != null;
                }
                else if (fileType == "excel")
                {
                    using var package = new ExcelPackage(new FileInfo(filePath));
                    check = package.Workbook.Worksheets.OfType<object>().Any<object>();
                }
            }
            catch
            {
                check = false;
            }

            return check;
        }

        static void ConfirmAndDeleteFiles(List<string> corruptedFiles)
        {
            Console.WriteLine("\nCorrupted files found:");
            foreach (var file in corruptedFiles)
            {
                Console.WriteLine(file);
            }

            Console.WriteLine("\nDo you want to delete these files? (y/n)");
            var response = Console.ReadLine();
            if (response?.ToLower() == "y")
            {
                foreach (var file in corruptedFiles)
                {
                    DeleteFile(file);
                }
            }
        }

        static void DeleteFile(string path)
        {
            try
            {
                File.Delete(path);
                Console.WriteLine($"Deleted file: {path}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error deleting file: {path}. Error: {ex.Message}");
            }
        }
    }
}
