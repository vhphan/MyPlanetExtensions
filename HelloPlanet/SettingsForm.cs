using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace HelloPlanet
{
    public partial class SettingsForm : Form
    {
        
        public SettingsForm()
        {
            InitializeComponent();
            textBox1.Text = Settings.ExcelToDoFile;
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
                Settings.ExcelToDoFile = openFileDialog1.FileName;
            }  
        }

        private void SettingsForm_Load(object sender, EventArgs e)
        {
        }


        private void SettingsForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            File.WriteAllText(Settings.ConfigFilePathJson, Utils.ObjToJSON(new Settings()));
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Load config.json file
            Settings settings = JsonConvert.DeserializeObject<Settings>(File.ReadAllText(Settings.ConfigFilePathJson));
            
            // Load values to form  
            textBox2.Text = Settings.ExcelToDoFile;
            textBox3.Text = Settings.ExcelExportFolder;
            
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
                Settings.ConfigFilePathJson = openFileDialog1.FileName;
                textBox1.Text = openFileDialog1.FileName;
            }
        }


    }
}