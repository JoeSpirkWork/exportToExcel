using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Bentley.Interop.MicroStationDGN;
using Bentley.DgnPlatformNET;

namespace exportToExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        bentleyCustomTools _bentleyTools;

        //This box will be used to name the Excel File
        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        //Asks the User to choose a folder the Excel File will be saved to
        private void button1_Click(object sender, EventArgs e)
        {
            
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if(result == DialogResult.OK)
            {
                folderBox.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        //This button will be utilized to ask the user to select items in microstation drawing. 
        private void button3_Click(object sender, EventArgs e)
        {
            _bentleyTools.selectElement();
            
        }


        private void button2_Click(object sender, EventArgs e)
        {
            Form Form1 = Form.ActiveForm;
            Form1.Close();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            if (folderBox.Text.Length > 0 && fileNameBox.Text.Length > 0)
            {
                excelFileTools excelFile = new excelFileTools();
                bentleyCustomTools bentleyTools = new bentleyCustomTools();
                _bentleyTools = bentleyTools;
                excelFile.createExcelFile(folderBox.Text, fileNameBox.Text);

                bentleyTools._excelFileRecord = excelFile;
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Ensure the folder and file box are filled in");
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            //allows the user to select an existing excel file
            DialogResult result = openFileDialog1.ShowDialog();

            if(result == DialogResult.OK)
            {
                //get the directory and name of the file from the open file dialog
                string fileDirectoryAndName = openFileDialog1.FileName;

                //create a new instance of the excel file tools object
                excelFileTools xlft = new excelFileTools();

                //sends the excel file to excel files tools for processing
                xlft.WorkWithExistingFile(fileDirectoryAndName);

                bentleyCustomTools bentleyTools = new bentleyCustomTools();
                _bentleyTools = bentleyTools;

                bentleyTools._excelFileRecord = xlft;
;
            }
            else
            {
                this.Close();
            }
        }
    }
}
