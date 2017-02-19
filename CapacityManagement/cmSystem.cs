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
                //Make Ericsson cell names all uppercase
                for (int i = 0; i < cellNamesEricsson.Length; i++)
                {
                    cellNamesEricsson[i].cellName = cellNamesEricsson[i].cellName.ToUpper();
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

        /*********
        *   Function:   cmCellsOvershootingUpdate
        *   Parameters: cmCell array of cells to update their overshooting value and recommended down tilt value
        *   Returns:    cmCell array of updated cells
        *   Operation Summary:
        *       Read overShooting xlsx file, clean data from the xlsx file and update the given array of cmCell[]
        *********/
        public static cmCell[] cmCellsOvershootingUpdate(cmCell[] cells)
        {
            //Read the overshooter files
            var overShootingCellsNamesFile = new ExcelQueryFactory("excelFiles/overShooters.xlsx");
            overShootingCellsNamesFile.ReadOnly = true;
            overShootingCellsNamesFile.UsePersistentConnection = true;
            overShootingCellsNamesFile.AddMapping<cmCell>(x => x.cellName, "CellName");
            overShootingCellsNamesFile.AddMapping<cmCell>(x => x.parameters.overShooting, "IsOvershoot");
            overShootingCellsNamesFile.AddMapping<cmCell>(x => x.recommendations.antennaDownTilt, "Down tilt");
            try
            {
                //We use with noHeader to ensure it reads tables as string and to avoid guessing it to be integers instead
                var overShootingCellNamesEricsson = (from x in overShootingCellsNamesFile.WorksheetNoHeader("ER Overshooters")
                                                     select x).ToArray();
                var overShootingCellNamesHuawei = (from x in overShootingCellsNamesFile.WorksheetNoHeader("HU Overshooters")
                                                   select x).ToArray();
                //Ericsson Data Cleaning and updating cell array
                for (int i = 1; i < overShootingCellNamesEricsson.Length; i++)
                {
                    //Data Cleaning
                    string cellName = overShootingCellNamesEricsson[i][0].ToString().ToUpper();
                    if (cellName.IndexOf("M", StringComparison.CurrentCultureIgnoreCase) != 0)
                    {
                        if (!Char.IsLetter(cellName.ElementAt(0)))
                        {
                            string str = "";
                            int len = cellName.Length;
                            for (int m = len; m < 5; m++)
                            {
                                str = "0" + str;
                            }
                            cellName = "G" + str + cellName;
                        }
                        else
                        {
                            cellName = "G" + cellName;
                        }
                    }
                    //Updaing cell array of Ericsson
                    for (int m = 0; m < cells.Length; m++)
                    {
                        //Console.WriteLine(cells[m].cellName == cellName);
                        if ((cells[m].cellName == cellName) && (cells[m].cellVendor == cellVendorType.Ericsson))
                        {
                            cells[m].parameters.overShooting = true;
                            cells[m].recommendations.antennaDownTilt = int.Parse(overShootingCellNamesEricsson[i][4]);
                            //Console.WriteLine(cells[m].parameters.overShooting + " " + cells[m].recommendations.antennaDownTilt);
                            break;
                        }
                    }
                }
                //Huawei Data Cleaning
                for (int i = 1; i < overShootingCellNamesHuawei.Length; i++)
                {
                    string cellName = overShootingCellNamesHuawei[i][0].ToString().ToUpper();
                    /* Case : Founding K in First */
                    if (cellName.IndexOf("K", StringComparison.CurrentCultureIgnoreCase) != -1)
                    {
                        cellName = 'G' + cellName;
                    }
                    /* Case : Founding Z in First */
                    else if (cellName.IndexOf("Z", StringComparison.CurrentCultureIgnoreCase) != -1)
                    {
                        cellName = 'G' + cellName;
                    }
                    /* Case :  ckecking the length of numbers  , Adding zeros and G */
                    else
                    {
                        int d = cellName.Length;
                        while (d < 5)
                        {
                            cellName = '0' + cellName;
                            d++;
                        }
                        cellName = 'G' + cellName;
                    }
                    //Updating Huawei cells with new overshooting values
                    for (int m = 0; m < cells.Length; m++)
                    {
                        if ((cells[m].cellName == cellName) && (cells[m].cellVendor == cellVendorType.Huawei))
                        {
                            cells[m].parameters.overShooting = true;
                            cells[m].recommendations.antennaDownTilt = int.Parse(overShootingCellNamesHuawei[i][4]);
                            break;
                        }
                    }
                }
            }
            finally
            {
                //Close connection
                overShootingCellsNamesFile.Dispose();
            }
            return cells;
        }
    }
}
