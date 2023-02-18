using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xl = Microsoft.Office.Interop.Excel;
using Microsoft.CSharp.RuntimeBinder;
using MessageBox = System.Windows.Forms.MessageBox;

namespace exportToExcel
{

    class excelFileTools
    {
        public excelFileTools() { }

        //The global parameters in the excel file tools
        public xl.Application _excel;
        public xl.Workbook _workBook;
        public xl.Worksheet _workSheet;
        public xl.Range _cellRange;

        //Creates a new excel file for data extraction
        public void createExcelFile(string folderName, string fileName)








        {
            xl.Application excel = null;
            xl.Workbook workBook = null;
            xl.Worksheet workSheet = null;
            xl.Range cellRange = null;


            try
            {
                //Creates a new app, file, worksheet, and range, so we have something to work with
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = true;
                excel.DisplayAlerts = true;

                workBook = excel.Workbooks.Add(Type.Missing);

                workSheet = workBook.ActiveSheet;
                workSheet.Name = "Civil Objects";

                //cellRange = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[1, 17]].Merge();


                //Names each of the new columns as per what we want to pull from Bentley

                string[] colNames = new string[17] { "ElemType", "BegX", "BegY", "BegZ", "EndX", "EndY", "EndZ", "Length", "Area", "CenterX", "CenterY", "CenterZ", "Rotation Degrees", "Text or Cell", "Named Group", "Level Name", "Element ID" };

                for (int Y = 0; Y < 17; Y++)
                {
                    workSheet.Cells[1, Y + 1] = colNames[Y];
                }

                //Save the file to where we want it to be, as the name we want it to be
                string saveExcelHere = folderName + "\\" + fileName;

                workBook.SaveAs(saveExcelHere);

                _excel = excel;
                _workBook = workBook;
                _workSheet = workSheet;


            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                excel.Quit();
                if (excel != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(excel); }
                if (workBook != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(workBook); }
                if (workSheet != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(workSheet); }
            }
            finally
            {



            }

        }

        //Writes data from Openroads/Microstation to Excel file on worksheet titled "extract"
        public void writeToExcel(object[] dataRecord)
        {

            xl.Application excel = _excel;
            xl.Workbook workBook = _workBook;
            xl.Worksheet workSheet = _workSheet;


            try
            {

                //cellRange = workSheet.Cells[workSheet.Cells[1,1],workSheet.Cells[1, 17] as xl.Range];
                int row = nextBlankRow(workSheet);
                //cellRange = workSheet.Range[workSheet.Cells[row, 1], workSheet.Cells[row, 17]];

                int cycle = 0;
                foreach (var properties in dataRecord)
                {
                    workSheet.Cells[row, cycle + 1] = dataRecord[cycle];
                    cycle++;
                }
                workBook.Save();

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                excel.Quit();
                if (excel != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(excel); }
                if (workBook != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(workBook); }
                if (workSheet != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(workSheet); }
            }

        }

        //finds the first blank row in worksheet and returns the row number
        private static int nextBlankRow(xl.Worksheet worksheet)
        {
            int i = 1;
            while ((worksheet.Cells[i, 1]).Value != null)
            {
                i++;
            }
            return i;

        }

        //Takes an existing excel file, ensures it has a title called "extract" and sets it to the global parameters
        public void WorkWithExistingFile(string folderAndFileName)
        {
            //Sets the 4 things we need to work with an excel file
            xl.Application excel = null;
            xl.Workbook workBook = null;
            xl.Worksheet workSheet = null;
            xl.Range cellRange = null;
            //Sets the boolean for determining whether the worksheet exists vs. whether we have to create it. 
            bool found = false;

            try
            {
                excel = new xl.Application();
                excel.Visible = true;
                workBook = excel.Workbooks.Open(folderAndFileName);
                
                foreach (xl.Worksheet workSheetJawn in workBook.Sheets)
                {
                    if(workSheetJawn.Name == "extract")
                    {
                        found = true;
                        break;
                    }
                }

                if (found)
                {
                    workSheet = workBook.Sheets["extract"];
                }
                else
                {
                    createWorkSheet(excel, workBook);
                    workSheet = workBook.Sheets["extract"];
                }

                //sets the global family name parameters to the existing excel file
                _excel = excel;
                _workBook = workBook;
                _workSheet = workSheet;

            }
            catch
            {
                MessageBox.Show("Darn, it's not working");
                excel.Quit();
                if (excel != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(excel); }
                if (workBook != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(workBook); }
                if (workSheet != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(workSheet); }
            }

        }
        
        //Creates a worksheet called "extract" with the correct parameters if the existing file chosen doesn't have 1
        public void createWorkSheet(xl.Application Excel, xl.Workbook workBook)
        {
            xl.Worksheet workSheet = workBook.Worksheets.Add();
            workSheet.Name = "extract";

            string[] colNames = new string[17] { "ElemType", "BegX", "BegY", "BegZ", "EndX", "EndY", "EndZ", "Length", "Area", "CenterX", "CenterY", "CenterZ", "Rotation Degrees", "Text or Cell", "Named Group", "Level Name", "Element ID" };

            for (int Y = 0; Y < 17; Y++)
            {
                workSheet.Cells[1, Y + 1] = colNames[Y];
            }

            workBook.Save();

        }
    }
}
