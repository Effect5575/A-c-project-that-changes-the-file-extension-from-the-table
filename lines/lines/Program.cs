using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace lines
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Okunacak excel dosya yolunu giriniz: ");
                string filePath = Console.ReadLine();
                Console.WriteLine("Aranacak dosya yolunu giriniz: ");
                string filePathSearch = Console.ReadLine(); string searchFileName = "";
                string fileUpdate = "";
                /*string filePath = @"C:\SearchFilesInExcel\Files.xlsx";
                string filePathSearch = @"C:\SearchFilesInExcel\TestFolder";*/
                List<string> filesToSearch = ReadExcelFile(filePath);
                List<string> foundFiles = GetAllFolders(filePathSearch); foreach (var foundFile in foundFiles)
                {
                    var ch = "\\";
                    int startNumber = foundFile.LastIndexOf(ch) + 1;
                    searchFileName = foundFile.Substring(startNumber); string newFileName = filesToSearch.FirstOrDefault(s => s.Contains(searchFileName)); if (!String.IsNullOrEmpty(newFileName))
                    {
                        fileUpdate = foundFile.Substring(0, startNumber) + newFileName;
                        File.Move(foundFile, fileUpdate);
                        Console.WriteLine("Güncellenen Dosya: {0}", fileUpdate);
                    }
                    newFileName = "";
                }                 /*foreach (var matchedFile in matchedFiles)
                {
                    Console.WriteLine("Bulunan Dosya: {0}", matchedFile);
                }*/
            }
            catch (Exception ex)
            {
                Console.WriteLine("Hata: {0}", ex.Message);
            }
            Console.WriteLine("Programı kapatmak için bir tuşa basınız!");
            Console.ReadKey();
        }
        public static List<string> ReadExcelFile(string _filePath)
        {
            List<string> Response = new List<string>();
            IExcelDataReader excelReader;
            int counter = 0; FileStream stream = File.Open(_filePath, FileMode.Open, FileAccess.Read);
            excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream); while (excelReader.Read())
            {
                counter++; if (counter > 1)
                {
                    Response.Add(excelReader.GetString(0) + "." + excelReader.GetString(1));
                }
            }
            excelReader.Close(); return Response;
        }
        public static List<string> GetAllFolders(string _searchFilePath)
        {
            List<string> Response = new List<string>();
            DirectoryInfo Directory = new DirectoryInfo(_searchFilePath);
            FileInfo[] files = Directory.GetFiles(".", SearchOption.AllDirectories); foreach (FileInfo file in files)
            {
                Response.Add(file.FullName);
            }
            return Response;
        }
    }
}
