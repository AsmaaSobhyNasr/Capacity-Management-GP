using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CapacityManagement
{
    public class cmCell
    {
        public string cellName { get; set; }
        public string cellType { get; set; }
        public struct recommendationsType
        {
            float antennaDownTilt;
        }
        public recommendationsType recommendations = new recommendationsType();

        public cmCell()
        {

        }
        public bool isCellOverShooting()
        {

            return true;
        }
    }
}
