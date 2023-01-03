using System;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using MapInfo.Types;
using Excel = Microsoft.Office.Interop.Excel;
using MapInfo.MiPro;
using MapInfo.MiPro.Interop;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

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

            foreach (var sheetName in GetExcelSheetNames(connectionString))
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
            con = new OleDbConnection(connectionString);
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

        public static void MakeConfigFile()
        {
            //C:\Users\vhphan\AppData\Local\InfoVista\Planet\7.7\Planet Extensions\HelloPlanet
            string appDataPath =
                Environment.GetFolderPath(Environment.SpecialFolder
                    .LocalApplicationData); // // C:\Users\vhphan\AppData\Local
            string appFolder = Path.Combine(appDataPath, "InfoVista", "Planet", "7.7", "Planet Extensions",
                "HelloPlanet"); // C:\Users\vhphan\AppData\Local\InfoVista\Planet\7.7\Planet Extensions\HelloPlanet
            if (!Directory.Exists(appFolder))
            {
                Directory.CreateDirectory(appFolder);
            }

            // Create the config.txt file in appFolder if it doesn't exist
            var configPath = Path.Combine(appFolder, "config.txt");
            var configPathJson = Path.Combine(appFolder, "config.json");    
            if (!File.Exists(configPath))
            {
                var configFile =  File.Create(configPath);
                configFile.Close();
            }
            if (!File.Exists(configPathJson))
            {
                var configFileJson =  File.Create(configPathJson);
                configFileJson.Close();
            }
            

            // Check Planet version
            Debug.WriteLine(System.Configuration.ConfigurationManager.AppSettings["PlanetVersion"]);
            Debug.WriteLine(configPath);
            Debug.WriteLine(Application
                .LocalUserAppDataPath); //  "C:\\Users\\vhphan\\AppData\\Local\\Precisely\\MapInfo Pro\\21.0.1.0025"
            Settings.ConfigFilePath = configPath;
            Settings.ConfigFilePathJson = configPathJson;
            
            // Create folder for storing exported files for Excel
            var excelFolder = Path.Combine(appFolder, "ExcelExport");
            if (!Directory.Exists(excelFolder))
            {
                Directory.CreateDirectory(excelFolder);
            }
            Settings.ExcelExportFolder = excelFolder;
            
            // Create folder for storing exported files for Planet
            var planetFolder = Path.Combine(appFolder, "PlanetExport");
            if (!Directory.Exists(planetFolder))
            {
                Directory.CreateDirectory(planetFolder);
            }
            Settings.PlanetExportFolder = planetFolder;
            
            
            
        }
        public static JObject MapStaticClassToJson(Type staticClassToMap)
        {
            var result = new JObject();
            var properties = staticClassToMap.GetProperties(BindingFlags.Public | BindingFlags.Static);
            foreach (PropertyInfo prop in properties)
            {
                result.Add(new JProperty(prop.Name, prop.GetValue(null, null)));
            }
    
            return result;
        }

        public static string ObjToJSON(object Obj)
        {
            return JsonConvert.SerializeObject(Obj, Formatting.Indented);
        }
    }
}