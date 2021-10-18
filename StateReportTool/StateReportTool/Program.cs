using System;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace StateReportTool
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("-----Processing State Impact Report-----");

            string filePath = ConfigurationManager.AppSettings["filePath"];           
            
            DataTable mainTable = new DataTable();
            DataTable workBookTable = new DataTable();

            DirectoryInfo dir = new DirectoryInfo(filePath);
            FileInfo[] files = dir.GetFiles("*.txt");
            if (files.Length != 0)
            {
                string[] lines = File.ReadAllLines(filePath + "\\" + files[0].Name);
                //lines [0] has column names
                //Split all the column names
                string[] cols = lines[0].Split('\t');                
                foreach (var colName in cols)
                {
                    mainTable.Columns.Add(new DataColumn(colName.ToString()));
                }
            }
           
            using (OleDbConnection conn = new OleDbConnection())
            {
                foreach (FileInfo file in files)
                {
                    DataTable dt = mainTable.Clone();
                    string[] lines = File.ReadAllLines(filePath + "\\" + file.Name);
                    List<string> list = new List<string>(lines);                    
                    list.RemoveAt(0);
                    lines = list.ToArray();
                    foreach (string line in lines)
                    {
                        var cols = line.Split('\t');

                        DataRow dr = dt.NewRow();
                        for (int cIndex = 0; cIndex < cols.Length; cIndex++)
                        {
                            dr[cIndex] = cols[cIndex];
                        }

                        dt.Rows.Add(dr);
                    }
                 
                    mainTable.Merge(dt);
                }
                
                var query = from row in mainTable.AsEnumerable()
                            group row by row.Field<string>("recip_address_state") into stateReport
                            orderby stateReport.Count() descending
                            select new
                            {
                                State = stateReport.Key,
                                Count = stateReport.Count()
                            };

                workBookTable.Columns.Add(new DataColumn("State"));
                workBookTable.Columns.Add(new DataColumn("Count"));
                // create a data table for WorkBook
                foreach (var obj in query)
                {
                    DataRow dr = workBookTable.NewRow();
                    dr["State"] = obj.State;
                    dr["Count"] = obj.Count;
                    workBookTable.Rows.Add(dr);
                }

                ExportToExcel(workBookTable);                
                Console.ReadKey();
            }

        }

        static void ExportToExcel(DataTable dataTable)
        {            
            object misValue = System.Reflection.Missing.Value;
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook  = xlApp.Workbooks.Add(misValue);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            for (int i = 1; i < dataTable.Columns.Count + 1; i++)
            {
                xlWorkSheet.Cells[1, i] = dataTable.Columns[i - 1].ColumnName;
            }
            for (int j = 0; j < dataTable.Rows.Count; j++)
            {
                for (int k = 0; k < dataTable.Columns.Count; k++)
                {
                    xlWorkSheet.Cells[j + 2, k + 1] = dataTable.Rows[j].ItemArray[k].ToString();
                }
            }          

            xlWorkBook.SaveAs(ConfigurationManager.AppSettings["excelPath"], Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            Console.WriteLine("State Impact report generated at " +  ConfigurationManager.AppSettings["excelPath"]);
        }

        static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
               Console.WriteLine("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
