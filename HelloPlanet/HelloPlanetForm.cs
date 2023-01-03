using System.Collections.Generic;
using System;
using System.Diagnostics;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Marconi.Wnp.Core;
using Marconi.Wnp.RFServices.Analysis;
using PlanetOpenApi.Analysis;
using PlanetOpenApi.Services;

namespace HelloPlanet
{
    public partial class HelloPlanetForm : Form
    {
        public HelloPlanetForm()
        {
            InitializeComponent();
        }

        readonly RunnerOfSorts _runner = new RunnerOfSorts();

        private void button1_Click(object sender, System.EventArgs e)
        {
            textBox1.Text = "Hello, Planet! Code 3488";
            textBox1.Text += Environment.NewLine;

            List<string> analysis = ServiceBase.AnalysisService.GetAnalysisNames(AnalysisType.LTEFDD);
            Debug.WriteLine(analysis.Count);
            foreach (string name in analysis)
            {
                Console.WriteLine(name);
                textBox1.Text += name + Environment.NewLine;
            }
        }

        public void updateTextBox2(string text)
        {
            textBox2.Text += text + Environment.NewLine;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            _runner.RunSomething(this);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            _runner.RunSomeMapinfoStuff(this);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            _runner.RunAnotherStuff(this);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Form settingsForm = new SettingsForm();
            settingsForm.Show();
            updateTextBox2(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile));
        }


        private void button6_Click(object sender, EventArgs e)
        {
            updateTextBox2(Settings.ConfigFilePath);
            updateTextBox2(Settings.ConfigFilePathJson);
            updateTextBox2(Settings.ExcelToDoFile);
        }

        private void HelloPlanetForm_Load(object sender, EventArgs e)
        {
            Utils.MakeConfigFile();
        }
    }
}