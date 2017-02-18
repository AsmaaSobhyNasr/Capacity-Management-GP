using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LinqToExcel;

namespace CapacityManagement
{
    public static class cmSystem
    {
        /*********
        *   Function:   cmGetCellNames
        *   Parameters: none
        *   Returns:    cmCell array of cells read from allCells.xlsx file
        *   Operation Summary:
        *       Read the xlsx file and edit Huawei cells to cellVendorType Huawei and then return the array
        *********/
        public static cmCell[] cmGetCellNames()
        {
            //Opening allCells.xlsx file with read-only mode and mapping to the cmCell Class
            var cellNamesFile = new ExcelQueryFactory("excelFiles/allCells.xlsx");
            cellNamesFile.ReadOnly = true;
            cellNamesFile.UsePersistentConnection = true;
            cellNamesFile.AddMapping<cmCell>(x => x.cellName, "Cell");
            cmCell[] cellNames;
            //the try and catch is for closing connection for allCells.xlsx file
            try
            {
                //Get the Huawei cell names from the file and change to array
                var cellNamesHuawei = (from x in cellNamesFile.Worksheet<cmCell>("HU Cells")
                                       where x.cellName != ""
                                       select x).ToArray<cmCell>();
                //Get the Ericsson cell names from the file and change to array
                var cellNamesEricsson = (from x in cellNamesFile.Worksheet<cmCell>("ER Cells")
                                         where x.cellName != ""
                                         select x).ToArray<cmCell>();
                //Change the vendor type for Huawei cells to Huawei as the default vendor type for no-parameters cmCell constructor is Ericsson
                for (int i = 0; i < cellNamesHuawei.Length; i++)
                {
                    cellNamesHuawei[i].cellVendor = cellVendorType.Huawei;
                }
                //Concatinate the two arrays
                cellNames = new cmCell[cellNamesEricsson.Length + cellNamesHuawei.Length];
                cellNamesEricsson.CopyTo(cellNames, 0);
                cellNamesHuawei.CopyTo(cellNames, cellNamesEricsson.Length);
            }
            finally
            {
                //Close connection
                cellNamesFile.Dispose();
            }
            return cellNames;
        }
        public static void cmCellsOvershootingUpdate(cmCell[] cells)
        {

        }
        public static void cmUpdateCells(cmCell cellNames, cmCell overShooting)
        {

        }
    }
}
