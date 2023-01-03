using Newtonsoft.Json;

namespace HelloPlanet
{
    public class Settings
    {
        public static string ConfigFilePath;
        public static string ConfigFilePathJson;
        [JsonProperty]
        public static string ExcelToDoFile = "";
        [JsonProperty]
        public static string ExcelExportFolder = "";
        [JsonProperty]
        public static string PlanetExportFolder = "";
        [JsonProperty]
        public static string ExcelTemplateFolder = "";
        [JsonProperty]
        public static string BestServerPath = "";
        [JsonProperty]
        public static string WestMalaysiaProjectTemplate = "";
        [JsonProperty]
        public static string EastMalaysiaProjectTemplate = "";

        // method to convert class to json string
        public static string ToJson()
        {
            return JsonConvert.SerializeObject(new Settings(), Formatting.Indented);
        }

    }
}