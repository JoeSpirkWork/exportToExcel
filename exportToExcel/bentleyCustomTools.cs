using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using proof = System.Windows.Forms;
using BDNET = Bentley.DgnPlatformNET;
using BDELEM = Bentley.DgnPlatformNET.Elements;
using Microstation = Bentley.MstnPlatformNET;
using BIM = Bentley.Interop.MicroStationDGN;
using BGeomNet = Bentley.GeometryNET;

namespace exportToExcel
{
    
    //This class compiles all of the resources, names, and functions, that we will use to interact with Microstation
    //JJS

    class bentleyCustomTools
    {
        public bentleyCustomTools() { }

        BDNET.DgnModelRef _activeModel;
        BDNET.DgnFile _activeFile;
        Bentley.Interop.MicroStationDGN.Application _app;

        public excelFileTools _excelFileRecord;

        public void selectElement()
        {
            //Gets the Active Model and file We're working in. 
            BDNET.DgnFile activeFile = Microstation.Session.Instance.GetActiveDgnFile();
            BDNET.DgnModelRef activeModel = Microstation.Session.Instance.GetActiveDgnModelRef();

            //Gets the App we're working in so we can use App Tools
            BIM.Application app = Bentley.MstnPlatformNET.InteropServices.Utilities.ComApp;

            //sets the global variables
            _activeFile = activeFile;
            _activeModel = activeModel;
            _app = app;

            //At this point, the user should have a selection. We will take the current selection and for now, show the beginning point. 
            uint numSelected = BDNET.SelectionSetManager.NumSelected();
            try
            {
                //Hashtable namedGroupDict = new Hashtable();
                //BDNET.NamedGroupCollection namedGroupCollection = BIM.NamedGroupMember

                for (uint i = 0; i < numSelected; i++)
                {
                    //Initiates a null, then runs a status for each element within the selection set. 
                    BDNET.Elements.Element el = null;
                    var status = BDNET.SelectionSetManager.GetElement(i, ref el, ref activeModel);

           
                    //This is will sort each element. My task for tomorrow and over the weekend is to create seperate functions that are called by picking out each element type. 
                   
                    if (el.ElementType == BDNET.MSElementType.Line)
                    {
                        BIM.Element line;
                        line = app.ActiveModelReference.GetElementByID(el.ElementId);
                        if (BDNET.NamedGroup.AnyGroupContains(el))
                        {
                            BDELEM.Element[] namedGroup = BDNET.NamedGroup.GetGroupsContaining(el);
                            System.Windows.Forms.MessageBox.Show(namedGroup[1].ToString());
                        }
                        processLine(line);
                    }
                    else if(el.ElementType == BDNET.MSElementType.LineString)
                    {
                        BIM.Element lineString;
                        lineString = app.ActiveModelReference.GetElementByID(el.ElementId);
                        processLineString(lineString);
                    }
                    else if (el.ElementType == BDNET.MSElementType.Shape)
                    {
                        BIM.Element Shape;
                        Shape = app.ActiveModelReference.GetElementByID(el.ElementId);
                        processShape(Shape);
                    }
                    else if (el.ElementType == BDNET.MSElementType.ComplexShape)
                    {
                        BDELEM.ComplexShapeElement compElement;
                        var query = new 
                        
                        compElement = app.ActiveModelReference.GetElementByID(el.ElementId) as BDELEM.ComplexShapeElement;
                        processComplexShape(compElement);
                    }
                    else if (el.ElementType == BDNET.MSElementType.Arc)
                    {
                        BIM.ArcElement eArc;
                        eArc = app.ActiveModelReference.GetElementByID(el.ElementId) as BIM.ArcElement;
                        processEArc(eArc);
                    }
                    else if (el.ElementType == BDNET.MSElementType.ComplexString)
                    {
                        BIM.ComplexStringElement cChain;
                        cChain = app.ActiveModelReference.GetElementByID(el.ElementId) as BIM.ComplexStringElement;
                        processCChain(cChain);
                    }
                    else if (el.ElementType == BDNET.MSElementType.Ellipse)
                    {
                        BIM.EllipseElement ellipse;
                        ellipse = app.ActiveModelReference.GetElementByID(el.ElementId) as BIM.EllipseElement;
                        processEllipses(ellipse);
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("Email Joe.Spirk@rsandh.com and tell him what type of object needs to be added. Object Selected is: " + el.Description);
                    }
                }
            }
            catch(Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            

        }
        //Function to Process Regular Lines
        public void processLine(BIM.Element line)
        {
            object[] dataExport = new object[17];

            //This series of functions sets all of the parameters
            //Element type is set
            //lineRecord.ElemType = ElemType;
            dataExport[0] = "Line";


            //Beginning coordinates are set
            //lineRecord.BegX = line.AsLineElement().StartPoint.X;
            //lineRecord.BegY = line.AsLineElement().StartPoint.Y;
            //lineRecord.BegZ = line.AsLineElement().StartPoint.Z;

            dataExport[1] = line.AsLineElement().StartPoint.X;
            dataExport[2] = line.AsLineElement().StartPoint.Y;
            dataExport[3] = line.AsLineElement().StartPoint.Z;


            //Ending Coordinates are set
            //lineRecord.EndX = line.AsLineElement().EndPoint.X;
            //lineRecord.EndY = line.AsLineElement().EndPoint.Y;
            //lineRecord.EndZ = line.AsLineElement().EndPoint.Z;

            dataExport[4] = line.AsLineElement().EndPoint.X;
            dataExport[5] = line.AsLineElement().EndPoint.Y;
            dataExport[6] = line.AsLineElement().EndPoint.Z;

            //length and Area
            //lineRecord.Length = line.AsLineElement().Length;

            dataExport[7] = line.AsLineElement().Length;

            //lines do not have area. 

            dataExport[8] = 0;

            //Sets the Center Elements
            //lineRecord.CenterX = (lineRecord.BegX + lineRecord.EndX) / 2;
            //lineRecord.CenterY = (lineRecord.BegY + lineRecord.EndY) / 2;
            //lineRecord.CenterZ = (lineRecord.BegZ + lineRecord.EndZ) / 2;

            dataExport[9] = (line.AsLineElement().StartPoint.X + line.AsLineElement().EndPoint.X) / 2;
            dataExport[10] = (line.AsLineElement().StartPoint.Y + line.AsLineElement().EndPoint.Y) / 2;
            dataExport[11] = (line.AsLineElement().StartPoint.Z + line.AsLineElement().EndPoint.Z) / 2;

            //will need help on the rotation angle, TExt or Cell, and Named Group
            dataExport[12] = 0;
            dataExport[13] = 0;
            dataExport[14] = 0;


            //Level
            //lineRecord.LevelName = line.AsLineElement().Level.ToString();
            dataExport[15] = line.AsLineElement().Level.Name;

            //Element ID
            //lineRecord.ElementID = line.ID;
            dataExport[16] = line.ID;

            //once the record is created and filled in, write it to an excel file. 
            _excelFileRecord.writeToExcel(dataExport);




        }

        //Function to Process LineStrings
        public void processLineString(BIM.Element lineString)
        {
            object[] dataExport = new object[17];

            //This series of functions sets all of the parameters
            //Element type is set
            dataExport[0] = "LineString";


            //Beginning coordinates are set
            dataExport[1] = lineString.AsLineElement().StartPoint.X;
            dataExport[2] = lineString.AsLineElement().StartPoint.Y;
            dataExport[3] = lineString.AsLineElement().StartPoint.Z;


            //Ending Coordinates are set
            dataExport[4] = lineString.AsLineElement().EndPoint.X;
            dataExport[5] = lineString.AsLineElement().EndPoint.Y;
            dataExport[6] = lineString.AsLineElement().EndPoint.Z;

            //length and Area
            dataExport[7] = lineString.AsLineElement().Length;

            //lines do not have area. 

            dataExport[8] = 0;

            //Sets the Center Elements
            dataExport[9] = (lineString.AsLineElement().StartPoint.X + lineString.AsLineElement().EndPoint.X) / 2;
            dataExport[10] = (lineString.AsLineElement().StartPoint.Y + lineString.AsLineElement().EndPoint.Y) / 2;
            dataExport[11] = (lineString.AsLineElement().StartPoint.Z + lineString.AsLineElement().EndPoint.Z) / 2;

            //will need help on the rotation angle, TExt or Cell, and Named Group
            dataExport[12] = 0;
            dataExport[13] = 0;
            dataExport[14] = 0;


            //Level
            dataExport[15] = lineString.AsLineElement().Level.Name;

            //Element ID
            dataExport[16] = lineString.ID;

            //once the record is created and filled in, write it to an excel file. 
            _excelFileRecord.writeToExcel(dataExport);




        }

        //Function to Process Shapes
        public void processShape(BIM.Element shape)
        {
            object[] dataExport = new object[17];

            //This series of functions sets all of the parameters
            //Element type is set
            dataExport[0] = "Shape";


            //Beginning coordinates are set. Shapes do not have start points.
            dataExport[1] = 0;
            dataExport[2] = 0;
            dataExport[3] = 0;


            //Ending Coordinates are set. Shapes do not have end points.
            dataExport[4] = 0;
            dataExport[5] = 0;
            dataExport[6] = 0;

            //length and Area
            //dataExport[7] = shape.AsLineElement().Length;

            //lines do not have area. 

            dataExport[8] = shape.AsClosedElement().Area();

            //Sets the Center Elements
            dataExport[9] = shape.AsClosedElement().Centroid().X;
            dataExport[10] = shape.AsClosedElement().Centroid().Y;
            dataExport[11] = shape.AsClosedElement().Centroid().Z;

            //will need help on the rotation angle, TExt or Cell, and Named Group
            dataExport[12] = 0;
            dataExport[13] = 0;
            dataExport[14] = 0;


            //Level
            dataExport[15] = shape.AsShapeElement().Level.Name;

            //Element ID
            dataExport[16] = shape.ID;

            //once the record is created and filled in, write it to an excel file. 
            _excelFileRecord.writeToExcel(dataExport);


        }

        //Process Complex Shapes
        public void processComplexShape(BDELEM.ComplexShapeElement compElement)
        {
            object[] dataExport = new object[17];

            //This series of functions sets all of the parameters
            //Element type is set
            dataExport[0] = "Complex Shape";


            //Beginning coordinates are set. Shapes do not have start points.
            dataExport[1] = 0;
            dataExport[2] = 0;
            dataExport[3] = 0;


            //Ending Coordinates are set. Shapes do not have end points.
            dataExport[4] = 0;
            dataExport[5] = 0;
            dataExport[6] = 0;

            //length and Area
            //dataExport[7] = shape.AsLineElement().Length;

            //Area

            //BDNET.ElementGraphicsProcessor
            //dataExport[8] = compElement.AsAreaFillPropertiesEdit().;

            /*Sets the Center Elements
            dataExport[9] = compElement.Centroid().X;
            dataExport[10] = compElement.Centroid().Y;
            dataExport[11] = compElement.Centroid().Z;

            //will need help on the rotation angle, TExt or Cell, and Named Group
            dataExport[12] = 0;
            dataExport[13] = 0;
            dataExport[14] = 0;


            //Level
            dataExport[15] = compElement.AsShapeElement().Level.Name;

            //Element ID
            dataExport[16] = compElement.ID;

            //once the record is created and filled in, write it to an excel file. 
            _excelFileRecord.writeToExcel(dataExport);

            */
        }

        //Processes Arcs
        public void processEArc(BIM.ArcElement eArc)
        {
            object[] dataExport = new object[17];

            //This series of functions sets all of the parameters
            //Element type is set
            dataExport[0] = "Arc";

            //Beginning coordinates are set
            dataExport[1] = eArc.StartPoint.X;
            dataExport[2] = eArc.StartPoint.Y;
            dataExport[3] = eArc.StartPoint.Z;

            //Ending Coordinates are set. Shapes do not have end points.
            dataExport[4] = eArc.EndPoint.X;
            dataExport[5] = eArc.EndPoint.Y;
            dataExport[6] = eArc.EndPoint.Z;

            //length and Area
            dataExport[7] = eArc.Length;

            //lines do not have area. 

            dataExport[8] = eArc.AsClosedElement().Area();

            //Sets the Center Elements
            dataExport[9] = eArc.CenterPoint.X;
            dataExport[10] = eArc.CenterPoint.Y;
            dataExport[11] = eArc.CenterPoint.Z;

            //will need help on the rotation angle, TExt or Cell, and Named Group
            dataExport[12] = 0;
            dataExport[13] = 0;
            dataExport[14] = 0;


            //Level
            dataExport[15] = eArc.AsShapeElement().Level.Name;

            //Element ID
            dataExport[16] = eArc.ID;

            //once the record is created and filled in, write it to an excel file. 
            _excelFileRecord.writeToExcel(dataExport);


        }

        //Processes complex chains
        public void processCChain(BIM.ComplexStringElement cChain)
        {
            object[] dataExport = new object[17];

            //This series of functions sets all of the parameters
            //Element type is set
            dataExport[0] = "Complex Chain";

            //Beginning coordinates are set
            dataExport[1] = cChain.StartPoint.X;
            dataExport[2] = cChain.StartPoint.Y;
            dataExport[3] = cChain.StartPoint.Z;

            //Ending Coordinates are set. Shapes do not have end points.
            dataExport[4] = cChain.EndPoint.X;
            dataExport[5] = cChain.EndPoint.Y;
            dataExport[6] = cChain.EndPoint.Z;

            //length and Area
            dataExport[7] = cChain.Length;

            //lines do not have area. 

            dataExport[8] = 0;

            //Sets the Center Elements
            dataExport[9] = 0;
            dataExport[10] = 0;
            dataExport[11] = 0;

            //will need help on the rotation angle, TExt or Cell, and Named Group
            dataExport[12] = 0;
            dataExport[13] = 0;
            dataExport[14] = 0;


            //Level
            dataExport[15] = cChain.Level.Name;

            //Element ID
            dataExport[16] = cChain.ID;

            //once the record is created and filled in, write it to an excel file. 
            _excelFileRecord.writeToExcel(dataExport);


        }

        //Process Ellipses
        public void processEllipses(BIM.EllipseElement ellipse)
        {
            object[] dataExport = new object[17];

            //This series of functions sets all of the parameters
            //Element type is set
            dataExport[0] = "Ellipse";


            //Beginning coordinates are set. Shapes do not have start points.
            dataExport[1] = 0;
            dataExport[2] = 0;
            dataExport[3] = 0;


            //Ending Coordinates are set. Shapes do not have end points.
            dataExport[4] = 0;
            dataExport[5] = 0;
            dataExport[6] = 0;

            //length and Area
            //dataExport[7] = shape.AsLineElement().Length;

            //lines do not have area. 

            dataExport[8] = ellipse.AsClosedElement().Area();

            //Sets the Center Elements
            dataExport[9] = ellipse.AsClosedElement().Centroid().X;
            dataExport[10] = ellipse.AsClosedElement().Centroid().Y;
            dataExport[11] = ellipse.AsClosedElement().Centroid().Z;

            //will need help on the rotation angle, TExt or Cell, and Named Group
            dataExport[12] = 0;
            dataExport[13] = 0;
            dataExport[14] = 0;


            //Level
            dataExport[15] = ellipse.AsShapeElement().Level.Name;

            //Element ID
            dataExport[16] = ellipse.ID;

            //once the record is created and filled in, write it to an excel file. 
            _excelFileRecord.writeToExcel(dataExport);


        }

    }
}
