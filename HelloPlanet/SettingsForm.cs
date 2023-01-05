using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace HelloPlanet
{
    public partial class SettingsForm : Form
    {
        private Settings _settings;

        public SettingsForm()
        {
            InitializeComponent();
        }

        private void selectFileButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"C:\",
                Title = "Select Config File",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "xlsx",
                Filter = "xlsx files (*.xlsx)|*.xlsx",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
                _settings.ExcelToDoFile = openFileDialog1.FileName;
            }
        }

        private void SettingsForm_Load(object sender, EventArgs e)
        {
        }


        private void SettingsForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_settings == null) return;
            File.WriteAllText(textBox1.Text, Utils.ObjToJSON(_settings));
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Load config.json file
            _settings = JsonConvert.DeserializeObject<Settings>(File.ReadAllText(textBox1.Text));

            // Load values to form  
            if (_settings == null) return;
            textBox2.Text = _settings.ExcelToDoFile;
            textBox3.Text = _settings.ExcelExportFolder;
            textBox4.Text = _settings.PlanetExportFolder;
            textBox5.Text = _settings.ExcelTemplateFolder;
            textBox6.Text = _settings.BestServerPath;
            textBox7.Text = _settings.WestMalaysiaProjectTemplate;
            textBox8.Text = _settings.EastMalaysiaProjectTemplate;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Select config.json file
            var openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"C:\",
                Title = "Select Config JSON File",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "json",
                Filter = "json files (*.json)|*.json",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }


        private void button3_Click(object sender, EventArgs e)
        {
            PlanetTasks planetTasks = new PlanetTasks(_settings);
            planetTasks.ParseExcelFile();
            planetTasks.ProcessSites();
        }
    }
}