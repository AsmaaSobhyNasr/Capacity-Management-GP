using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CapacityManagement
{
    
    public class cmCell
    {
        //Properties
        public string cellName { get; set; }
        public cellVendorType cellVendor { get; set; }
        public struct parametersType
        {
            public bool overShooting { get; set; }
        }
        public parametersType parameters;
        public struct recommendationsType
        {
            public float antennaDownTilt;
        }
        public recommendationsType recommendations;

        //Constructors
        public cmCell()
        {
            this.cellName = "";
            this.cellVendor = cellVendorType.Ericsson;
            this.parameters.overShooting = false;
            this.recommendations.antennaDownTilt = 0;
        }
        public cmCell(string cellName, cellVendorType cellVendor)
        {
            this.cellName = cellName;
            this.cellVendor = cellVendor;
            this.parameters.overShooting = false;
            this.recommendations.antennaDownTilt = 0;
        }

        //Methods
        public bool isCellOverShooting()
        {
            return this.parameters.overShooting;
        }
    }
}
