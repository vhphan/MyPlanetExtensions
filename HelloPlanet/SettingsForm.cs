using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace HelloPlanet
{
    public partial class SettingsForm : Form
    {
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

                DefaultExt = "txt",
                Filter = "txt files (*.txt)|*.txt",
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

        private void SettingsForm_Load(object sender, EventArgs e)
        {
        }


        private void SettingsForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            
        }
    }
}