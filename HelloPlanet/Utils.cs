using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using MapInfo.Types;
using Excel = Microsoft.Office.Interop.Excel;


namespace HelloPlanet
{
    public static class Utils
    {
        public static Excel.Range ReadExcelFile(string filePath)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            
            CleanUp(xlApp);
            
            return xlRange;
        }

        private static void CleanUp(Excel.Application xlApp)
        {
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
        
        public static DataSet Parse(string fileName)
        {
            string connectionString =
                $"provider=Microsoft.ACE.OLEDB.16.0; data source={fileName};Extended Properties=Excel 8.0;";

            var data = new DataSet();

            foreach(var sheetName in GetExcelSheetNames(connectionString))
            {
                using (OleDbConnection con = new OleDbConnection(connectionString))
                {    
                    var dataTable = new DataTable();
                    string query = $"SELECT * FROM [{sheetName}]";
                    con.Open();
                    OleDbDataAdapter adapter = new OleDbDataAdapter(query, con);
                    adapter.Fill(dataTable);
                    data.Tables.Add(dataTable);
                }
            }

            return data;
        }

        static string[] GetExcelSheetNames(string connectionString)
        {
            OleDbConnection con = null;
            DataTable dt = null;
            con= new OleDbConnection(connectionString);
            con.Open();
            dt = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            if (dt == null)
            {
                return null;
            }

            String[] excelSheetNames = new String[dt.Rows.Count];
            int i = 0;

            foreach (DataRow row in dt.Rows)
            {
                excelSheetNames[i] = row["TABLE_NAME"].ToString();
                i++;
            }

            return excelSheetNames;
        }
        

        
    }
}