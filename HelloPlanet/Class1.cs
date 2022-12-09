using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using MapInfo.MiPro.Interop;
using MapInfo.Types;
using Marconi.Wnp.RFServices.Analysis;
using PlanetOpenApi.Analysis;
using PlanetOpenApi.Common.ProgressReporting;
using PlanetOpenApi.Data.Selection;
using PlanetOpenApi.Services;
using PlanetOpenApi.MapInfoLoader;
using PlanetOpenApi.Network;

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
            try
            {
                ServiceBase.AnalysisService.SetSectors(AnalysisType.LTEFDD, analysisName, sectors);
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
            MapBasicService.MapBasicDo("run application \"C:\\Users\\vhphan\\source\\repos\\HelloPlanet\\HelloPlanet\\HelloPy.py\"");
            
        }
    }
}