﻿
using System.Collections.Generic;
using System;
using System.Diagnostics;
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
            RunnerOfSorts runner = new RunnerOfSorts();
            runner.RunSomething(this);
        }
    }
}