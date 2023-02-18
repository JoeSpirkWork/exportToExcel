//--------------------------------------------------------------------------------------+
//
//    $Source: exportToExcel.cs $
// 
//    $Copyright: (c) RS&H 2022 $
//
//---------------------------------------------------------------------------------------+

namespace exportToExcel
{
    [Bentley.MstnPlatformNET.AddIn(MdlTaskID = "exportToExcel")]
    internal sealed class exportToExcel : Bentley.MstnPlatformNET.AddIn
    {
        //--------------------------------------------------------------------------------------
        // @description   This function does...
        // @bsimethod                                                    Bentley
        //+---------------+---------------+---------------+---------------+---------------+------
        private exportToExcel(System.IntPtr mdlDesc) : base(mdlDesc)
        {
        }
        //--------------------------------------------------------------------------------------
        // @description   This function does...
        // @bsimethod                                                    Bentley
        //+---------------+---------------+---------------+---------------+---------------+------
        protected override int Run(string[] commandLine)
        {
            return 0;
        }
    }



}