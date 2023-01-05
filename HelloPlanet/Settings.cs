using System.Collections.Generic;
using Newtonsoft.Json;

namespace HelloPlanet
{
    public class Settings
    {
        public string ConfigFilePathJson {get; set;}
        public string ExcelToDoFile { get; set; }
        public string ExcelExportFolder { get; set; }
        public string PlanetExportFolder {get; set;}
        public string ExcelTemplateFolder {get; set;}
        public string BestServerPath {get; set;}
        public string WestMalaysiaProjectTemplate {get; set;}
        public string EastMalaysiaProjectTemplate {get; set;}
        
        // method to convert class to json string
        public string ToJson()
        {
            return JsonConvert.SerializeObject(this, Formatting.Indented);
        }

        // method to return properties as a dictionary
        public static Dictionary<string, string> ToDictionary()
        {
            var dict = new Dictionary<string, string>();
            for (var index = 0; index < typeof(Settings).GetProperties().Length; index++)
            {
                var prop = typeof(Settings).GetProperties()[index];
                dict.Add(prop.Name, prop.GetValue(null).ToString());
            }

            return dict;
        }
    }
}