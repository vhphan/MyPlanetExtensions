using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;

namespace HelloPlanet
{
    public class PlanetTasks
    {

        private string[] _sitesList;
        private readonly string _excelFilePath;

        public PlanetTasks(Settings settings)
        {
            _excelFilePath = settings.ExcelToDoFile;
        }

        
        public void ParseExcelFile()
        {
            var result = Utils.Parse(_excelFilePath);
            //object value = DataSetObj.Tables["Table_Name"].Rows[rowIndex]["column_name"]
            foreach(DataRow row in result.Tables[0].Rows)
            {
                foreach (DataColumn column in result.Tables[0].Columns)
                {
                    Debug.WriteLine("====================================");
                    Debug.WriteLine("Column: " + column.ColumnName + " Value: " + row[column]);
                    Debug.WriteLine("====================================");
                }
            }
            _sitesList = result.Tables[0].AsEnumerable().Select(r => r.Field<string>("Site_ID")).ToArray();
        }
        
        public void ProcessSites()
        {
            foreach (var site in _sitesList)
            {
                // skip if site is null
                if (string.IsNullOrEmpty(site))
                    continue;
                ProcessSite(site);
            }
        }

        private void ProcessSite(string site)
        {
            // skip if site is null
            if (string.IsNullOrEmpty(site))
                return;
            
            Debug.WriteLine("====================================");
            Debug.WriteLine("Processing Site: " + site);
            Debug.WriteLine("====================================");

            
        }
    }
}