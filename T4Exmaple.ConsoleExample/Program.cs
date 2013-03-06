using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;

namespace T4Exmaple.ConsoleExample
{
    class Program
    {
        static void Main(string[] args)
        {
            //DirectoryInfo currentDir = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory);
            //Console.WriteLine();
            //string combineDir = Path.Combine(currentDir.Parent.Parent.FullName, "ResourceCSV");
            ////Console.WriteLine(combineDir);
            //if (Directory.Exists(combineDir))
            //{
            //    DirectoryInfo resourceDir = new DirectoryInfo(combineDir);
            //    FileInfo[] files = resourceDir.GetFiles("*.xml", SearchOption.AllDirectories);
            //    foreach (FileInfo item in files)
            //    {
            //        Console.WriteLine(item.Directory.Parent.Name);
            //        Console.WriteLine(item.Name.Replace(".xml",string.Empty));
            //        Console.WriteLine(item.FullName);
            //    }
            //}
            //else
            //{
            //    Console.WriteLine("Folder doesn't exist");
            //}

            CsvDataProvider provider = new CsvDataProvider(new DirectoryInfo(@"..\").Parent.FullName);
            Console.WriteLine(provider.AllDataTable.Tables.Count);
            //DataTable table = provider.AllDataTable.Tables[1];
            foreach (DataTable item in provider.AllDataTable.Tables)
            {
                Console.WriteLine(item.TableName);
            }
            Console.ReadLine();
        }
    }
}
