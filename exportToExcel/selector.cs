using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Bentley.Interop.MicroStationDGN;
using finder = Bentley.Interop.MicroStationDGN.ILocateCommandEvents;

namespace exportToExcel
{
    class selector : finder
    {
        Bentley.Interop.MicroStationDGN.Application app;

        public void Accept(Element Element, ref Point3d Point, View View)
        {
            throw new NotImplementedException();
        }

        public void LocateFailed()
        {
            throw new NotImplementedException();
        }

        public void LocateFilter(Element Element, ref Point3d Point, ref bool Accepted)
        {
            throw new NotImplementedException();
        }

        public void LocateReset()
        {
            throw new NotImplementedException();
        }

        public void Cleanup()
        {
            throw new NotImplementedException();
        }

        public void Start()
        {
           app = Bentley.MstnPlatformNET.InteropServices.Utilities.ComApp;

            app.ShowPrompt("Pick the Object you'd like to write to Excel");
        }

        public void Dynamics(ref Point3d Point, View View, MsdDrawingMode DrawMode)
        {
            throw new NotImplementedException();
        }
    }
}
