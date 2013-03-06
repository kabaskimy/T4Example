using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;

namespace DX.Framework.ResourceTemplate
{
    public class XlsDataProvider
    {
        public XlsDataProvider()
        {
        }

        public XlsDataProvider(string folderPath)
        {
            FolderPath = folderPath;
            FileNameList = new List<FileInfo>();
            AllDataTable = new DataSet();
            if (Directory.Exists(folderPath))
            {

                DirectoryInfo resourceDir = new DirectoryInfo(Path.Combine(folderPath, "ResourceCSV"));
                //Console.WriteLine(resourceDir.FullName);
                FileInfo[] files = resourceDir.GetFiles("*.xls?", SearchOption.AllDirectories);
                foreach (FileInfo item in files)
                {
                    FileNameList.Add(item);
                }
            }
            else
            {
                Console.WriteLine("Folder doesn't exist");
            }

            foreach (FileInfo item in FileNameList)
            {
                InitializeWorkbook(item);
                ConvertToDataTable(item);
            }
        }

        public string FolderPath
        {
            set;
            get;
        }

        public DataSet AllDataTable
        {
            private set;
            get;
        }

        public List<FileInfo> FileNameList
        {
            private set;
            get;
        }

        HSSFWorkbook hssfworkbook;
        private void InitializeWorkbook(FileInfo fileName)
        {
            //read the template via FileStream, it is suggested to use FileAccess.Read to prevent file lock.
            //book1.xls is an Excel-2007-generated file, so some new unknown BIFF records are added. 
            using (FileStream file = new FileStream(fileName.FullName, FileMode.Open, FileAccess.Read))
            {
                hssfworkbook = new HSSFWorkbook(file);
            }
        }

        private void ConvertToDataTable(FileInfo item)
        {
            ISheet sheet = hssfworkbook.GetSheetAt(0);
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();


            string tableFileName = item.Name.Substring(0, item.Name.Length - item.Extension.Length);
            //for (int j = 0; j < 5; j++)
            //{
            //    dt.Columns.Add(Convert.ToChar(((int)'A') + j).ToString());
            //}
            int columnNumber = 0;
            List<DataTable> multiTables = new List<DataTable>();
            if (rows.MoveNext())
            {
                IRow collumnRow = (HSSFRow)rows.Current;
                for (int i = 0; i < collumnRow.LastCellNum; i++)
                {
                    ICell collumnCell = collumnRow.GetCell(i);
                    if (collumnCell != null && !String.IsNullOrEmpty(collumnCell.ToString().Trim()))
                    {
                        //dt.Columns.Add(collumnCell.ToString().Trim());
                        if (i != 0)
                        {
                            DataTable multiData = new DataTable(tableFileName + "." + collumnCell.ToString().Trim());
                            multiData.Columns.Add("key");
                            multiData.Columns.Add(collumnCell.ToString().Trim());
                            multiTables.Add(multiData);
                        }
                        columnNumber++;
                    }
                    else
                    {
                        break;
                    }
                }

            }


            while (rows.MoveNext())
            {
                IRow row = (HSSFRow)rows.Current;
                //DataRow dr = dt.NewRow();

                string key = string.Empty;
                for (int i = 0; i < row.LastCellNum && i < columnNumber; i++)
                {
                    ICell cell = row.GetCell(i);
                    if (i == 0)
                    {
                        key = cell.ToString().Trim();
                    }
                    else
                    {
                        DataRow dr = multiTables[i - 1].NewRow();
                        dr[0] = key;
                        dr[1] = cell.ToString().Trim();
                        multiTables[i - 1].Rows.Add(dr);
                    }


                }

            }
            foreach (DataTable dt in multiTables)
            {
                AllDataTable.Tables.Add(dt);
            }
        }

    }
}
