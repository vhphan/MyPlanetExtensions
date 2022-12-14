using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using MapInfo.Data;
using Marconi.Wnp.RFServices.Analysis;
using PlanetOpenApi.Analysis;
using PlanetOpenApi.Common.ProgressReporting;
using PlanetOpenApi.Data.Selection;
using PlanetOpenApi.MapInfoLoader;
using PlanetOpenApi.Services;

namespace HelloPlanet
{
    internal class RunnerOfSorts
    {
        HelloPlanetForm _form;

        private void generationTask_WorkCompleted(object sender, WorkCompletedEventArgs e)
        {
            try
            {
                if (e?.Result == null)
                {
                    return;
                }

                this._form.updateTextBox2(e.Result.ToString());
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
            }
        }

        private void generationTask_ProgressChanged(object sender, ProgressEventArgs e)
        {
            try
            {
                if (e?.Message == null)
                {
                    return;
                }

                this._form.updateTextBox2(e.Message.ToString());
                Debug.WriteLine(e.Message.ToString());
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
            }
        }

        public void RunSomething(HelloPlanetForm sender)
        {
            // AnalysisService analysisService = AnalysisService;
            this._form = sender;
            String analysisName = "Analysis 1";

            List<SiteSectorId> sectors = new List<SiteSectorId>
            {
                new SiteSectorId("Crazy", "4"),
                new SiteSectorId("LTE_700MHz_1", "2"),
                new SiteSectorId("LTE_700MHz_1", "3"),
            };
            Debug.WriteLine("sectors: " + sectors.Count);
            try
            {
                ServiceBase.AnalysisService.SetSectors(AnalysisType.LTEFDD, analysisName, sectors);
                var sectors2 = ServiceBase.AnalysisService.GetSectors(AnalysisType.LTEFDD, analysisName);


                Debug.WriteLine("Sectors: " + sectors2.Count());
                foreach (var sector in sectors2)
                {
                    Debug.WriteLine(sector.SiteId);
                    Debug.WriteLine(sector.SectorId);
                    sender.updateTextBox2(sector.SiteId);
                    sender.updateTextBox2(sector.SectorId);
                }
                
            }
            catch (Exception e)
            {
                Debug.WriteLine(e);
                sender.updateTextBox2(e.ToString());
            }

            IProgressTask generationTask = AnalysisService.CreateAnalysisTask(AnalysisType.LTEFDD, analysisName, false);
            generationTask.WorkCompleted += generationTask_WorkCompleted;
            generationTask.ProgressChanged += generationTask_ProgressChanged;
            generationTask.RunAsync(null);
        }

        public void RunSomeMapinfoStuff(HelloPlanetForm sender)
        {
            this._form = sender;
            // Send mapbasic commands
            // InteropServices.MapInfoApplication.Do("set window 3004 hide");
            // MapBasicService.MapBasicDo("set window 3004 show");

            // Close the project
            // ProjectController.Close();

            // Export Network Data
            // Dictionary<String, ICollection<String>> tblInfo = new Dictionary<String, ICollection<String>>();
            // tblInfo.Add("Sites", new List<String> { "Site ID", "Longitude", "Latitude" });
            // ExportNetworkDataToCsvTask exportNetworkDataTask = new ExportNetworkDataToCsvTask(tblInfo, "C:\\Users\\vhphan\\temp");
            // exportNetworkDataTask.ProgressChanged += generationTask_ProgressChanged;
            // exportNetworkDataTask.WorkCompleted += generationTask_WorkCompleted;
            // exportNetworkDataTask.RunAsync(null);

            // Run Python Script from within Mapinfo Application
            MapBasicService.MapBasicDo(
                "run application \"C:\\Users\\vhphan\\source\\repos\\HelloPlanet\\HelloPlanet\\HelloPy.py\"");

            // Get current project path
            String projectFilePath = ProjectFolders.ProjectPath;
            sender.updateTextBox2(projectFilePath);

            // Query sites
            string query = @"Select * From SiteFile Where SiteName=""LTE_700MHz_17"" and Type=""Sector"" Into CheckSite";
            Debug.WriteLine(query);

            MIConnection connection = new MIConnection();
            MICommand command = null;

            try
            {
                connection.Open();
                command = connection.CreateCommand();
                command.CommandText = query;
                var count = (int)command.ExecuteScalar();
                command.Dispose();
                connection.Close();
                sender.updateTextBox2($"MISQL table record count: {count}");
            }
            catch (Exception ex)
            {
                // MessageBox.Show($"error: {ex.Message}");
                sender.updateTextBox2($"error: {ex.Message}");
            }
            finally
            {
                command?.Dispose();
                connection.Close();
            }
            sender.updateTextBox2("Finished running...");
            
        }

        public void RunAnotherStuff(HelloPlanetForm sender)
        {
            // C:\Users\vhphan\Documents\15 Site Data_20211030.xlsx
            string excelPath = System.Configuration.ConfigurationManager.AppSettings["SiteListExcelPath"];
            var result = Utils.Parse(excelPath);
            sender.updateTextBox2(result.ToString());
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
            
            string[] sites = result.Tables[0].AsEnumerable().Select(r => r.Field<string>("Site_ID")).ToArray();
            
        }
    }
}