using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
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
            List<string> filesToCheck = EnumerateAndFilterFiles(basePath);
            Console.WriteLine($"Number of files to check: {filesToCheck.Count}");

            if (filesToCheck.Count > 0)
            {
                List<string> corruptedFiles = SearchFiles(filesToCheck);
                if (corruptedFiles.Count > 0)
                {
                    ConfirmAndDeleteFiles(corruptedFiles);
                }
                else
                {
                    Console.WriteLine("No corrupted files found.");
                }
            }
            else
            {
                Console.WriteLine("No files found to check.");
            }
        }

        static List<string> EnumerateAndFilterFiles(string basePath)
        {
            var fileTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                ".png", ".jpg", ".jpeg",
                ".pdf",
                ".xls", ".xlsx", ".xlsm", ".xlsb", ".odf", ".ods", ".odt"
            };

            var files = Directory.EnumerateFiles(basePath, "*.*", SearchOption.AllDirectories)
                                 .Where(file => fileTypes.Contains(Path.GetExtension(file)))
                                 .ToList();

            return files;
        }

        static List<string> SearchFiles(List<string> files)
        {
            var corruptedFiles = new List<string>();

            foreach (var filePath in files)
            {
                Console.WriteLine(new string('-', 30));
                Console.WriteLine($"Checking file: {filePath}");

                string fileType = GetFileType(filePath);
                bool? check = CheckFile(filePath, fileType);

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

        static string GetFileType(string filePath)
        {
            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            return extension switch
            {
                ".png" or ".jpg" or ".jpeg" => "image",
                ".pdf" => "pdf",
                ".xls" or ".xlsx" or ".xlsm" or ".xlsb" or ".odf" or ".ods" or ".odt" => "excel",
                _ => null
            };
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
                    using var pdf = PdfReader.Open(filePath, PdfDocumentOpenMode.Import);
                    check = pdf.Info != null;
                }
                else if (fileType == "excel")
                {
                    using var package = new ExcelPackage(new FileInfo(filePath));
                    check = package.Workbook.Worksheets.OfType<object>().Any<object>();
                }
                else
                {
                    check = null;
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
