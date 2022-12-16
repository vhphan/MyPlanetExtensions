using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;

namespace HelloPlanet
{
    public partial class SettingsForm : Form
    {
        
        public SettingsForm()
        {
            InitializeComponent();
            textBox1.Text = Globals.excelToDoFile;
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
                Globals.excelToDoFile = openFileDialog1.FileName;
            }  
        }

        private void SettingsForm_Load(object sender, EventArgs e)
        {
        }


        private void SettingsForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // save _excelToDoPath to a text file
            // System.IO.File.WriteAllText(@"C:\Users\Public\Documents\HelloPlanet\Settings.txt", _excelToDoPath);
            // System.IO.File.WriteAllLines();  

            List<string> settings = new List<string>();

            if (File.Exists(Globals.excelToDoFile))
            {
                settings.Add("TODO=" + Globals.excelToDoFile);
            }

            File.WriteAllLines(Globals.configFilePath, settings);

        }
    }
}