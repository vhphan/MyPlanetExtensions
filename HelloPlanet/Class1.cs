using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Marconi.Wnp.RFServices.Analysis;
using PlanetOpenApi.Analysis;
using PlanetOpenApi.Common.ProgressReporting;
using PlanetOpenApi.Data.Selection;
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
                new SiteSectorId("LTE_700MHz_1", "1"),
                new SiteSectorId("LTE_700MHz_1", "2"),
                new SiteSectorId("LTE_700MHz_1", "3"),
            };
            ServiceBase.AnalysisService.SetSectors(AnalysisType.LTEFDD, analysisName, sectors);
            
            IProgressTask generationTask = AnalysisService.CreateAnalysisTask(AnalysisType.LTEFDD, analysisName, false);
            generationTask.WorkCompleted += generationTask_WorkCompleted;
            generationTask.ProgressChanged += generationTask_ProgressChanged;
            generationTask.RunAsync(null);
        }
    }
}