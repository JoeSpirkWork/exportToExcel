//--------------------------------------------------------------------------------------+
//
//    $Source: exportToExcelClass.cs $
// 
//    $Copyright: (c) 2022 Bentley Systems, Incorporated. All rights reserved. $
//
//---------------------------------------------------------------------------------------+

//---------------------------------------------------------------------------------------+
//	Using Directives
//---------------------------------------------------------------------------------------+
using System.Windows.Forms;
using BDNET = Bentley.Interop.MicroStationDGN;

namespace exportToExcel
{
    class exportToExcelClass
    {
        //--------------------------------------------------------------------------------------
        // @description   This function will create a c# form for the user to select the folder 
        //they would like to store the excel file, and then allows
        // the user to select multiple items from the microstation drawing they are working in. 
        // @bsimethod                                                    Bentley
        //+---------------+---------------+---------------+---------------+---------------+------


        public static void HelloWorld(string unparsed)
        {
            MessageBox.Show("Hello World");
        }


        //--------------------------------------------------------------------------------------
        // @This Method will begin the entire process. This function brings up the Form to begin the work
        // @JJS                                                   Bentley
        //+---------------+---------------+---------------+---------------+---------------+------
        public static void beginExport(string unparsed)
        {
            var Form1 = new Form1();
            Form1.Show();
        }


    }
}