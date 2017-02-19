using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LinqToExcel;

namespace CapacityManagement
{
    public enum cellVendorType { Ericsson, Huawei };
    class Program
    {
        static void Main(string[] args)
        {
            cmCell[] cellNamesArray = cmSystem.cmGetCellNames();
            cellNamesArray = cmSystem.cmCellsOvershootingUpdate(cellNamesArray);
            Console.ReadLine();

        }
    }
}
