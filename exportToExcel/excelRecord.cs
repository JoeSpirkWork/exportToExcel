using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace exportToExcel
{
    class excelRecord
    {
        //This defines all of our variables. 
        public String ElemType = null;
        public double BegX, BegY, BegZ, EndX, EndY, EndZ, Length, Area, CenterX, CenterY, CenterZ, RotationDegrees;
        public string NamedGroup, LevelName;
        public long ElementID;

        public enum TextorCell
        {
            Text,
            Cell
        }

        //Constructor - This will initiate a new instance of an excel record
        public excelRecord() { }

        //The following 17 functions allow the user to set the parameters of this object
        
        public void setElemType(string _ElemType)
        {
            this.ElemType = _ElemType;
        }

       

    }
}
