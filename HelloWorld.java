using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace PhoneNums
{
    internal class Program
    {
        static string dir;
        static bool isDirectoryFound = false;
        static async Task Main(string[] args)
        {
            try
            {
                while (!isDirectoryFound)
                {
                    Console.WriteLine("Ведите путь к файлу");
                    dir = Console.ReadLine();
                    if (Directory.Exists(dir))
                    {
                        Console.WriteLine("Директория найдена");
                        Console.WriteLine("Найденные файлы Excel:");
                        var directory = new DirectoryInfo(dir);
                        isDirectoryFound = true;
                        FileInfo[] files = directory.GetFiles("*.xlsx");
                        foreach (FileInfo s in files)
                        {
                        
                            Console.WriteLine(s);
                        }
                        
                        string fileName;
                        FileInfo currentFile = null;
                        bool isFileFound = false;
                        while (!isFileFound)
                        {
                            Console.WriteLine("Введите имя необходимого файлы");
                            fileName = Console.ReadLine() + ".xlsx";
                            foreach (FileInfo file in files)
                            {
                                if (file.Name == fileName)
                                {
                                    currentFile = file;
                                    isFileFound = true;
                                }
                            }
                        }
                        Console.WriteLine("Файл выбран");
                        Excel.Application application = new Excel.Application();
                        application.Workbooks.Open(currentFile.Directory.FullName + "\\" + currentFile.Name);
                        Excel.Worksheet sheet = (Excel.Worksheet)application.Worksheets.get_Item(1);
                        var lastCell = application.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                        string[,] list = new string[lastCell.Column, lastCell.Row];
                        for (int i = 0; i < lastCell.Column; i++)
                            for (int j = 0; j < lastCell.Row; j++)
                                list[i, j] = application.Cells[j + 1, i + 1].Text.ToString();
          
                        Regex regex = new Regex(@"^8\(\d{3}\)\d{3}-\d{2}-\d{2}");
                        List<string> phoneNumbers = new List<string>();
                        foreach (string str in list)
                        {
                            if (regex.Match(str).Success)
                            {
                                phoneNumbers.Add(str);
                            }
                        }

                        string path = currentFile.Directory.FullName + "\\phoneNumbers.txt";
                        if (!File.Exists(path))
                            File.Create(path);
                        bool first = false;
                        foreach (string str in phoneNumbers)
                        {
                            using (StreamWriter writer = new StreamWriter(path, first))
                            {
                                await writer.WriteLineAsync(str);
                            }
                            first = true;
                            Console.WriteLine(str);
                        }

                        application.Workbooks.Close();
                        application.Quit();
                        Console.WriteLine("Закрыто");
                        GC.Collect();
                    }
                }
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

    }
}