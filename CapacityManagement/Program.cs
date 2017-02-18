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
            Console.WriteLine(y);
            /*
            var cellNamesFile = new ExcelQueryFactory("excelFiles/allCells.xlsx");
            var overShootingCellsNamesFile = new ExcelQueryFactory("excelFiles/overShooters.xlsx");

            var cellNames = from x in cellNamesFile.WorksheetNoHeader("ER Cells")
                        select x;
            var overShootingCellNames = from k in overShootingCellsNamesFile.WorksheetNoHeader("ER Overshooters")
                                    select k;
            //Add cells to an array
            string[] arrayCell = new string[cellNames.Count()];
            try
            {
                int x = 0;
                foreach (var cell in cellNames)
                {
                    arrayCell[x] = cell[0].ToString().ToUpper();
                    x++;
                }
            }
            catch(System.ArgumentException e){
               
            }
            //Add overshooters to an array with recommendedTowntilt
            string[][] arrayOverShooters = new string[overShootingCellNames.Count()][];
            try
            {
                int x = 0;
                foreach (var overshooter in overShootingCellNames)
                {
                    arrayOverShooters[x] = new string[] { overshooter[0].ToString().ToUpper(), overshooter[4].ToString() };
                    x++;
                }

            }
            catch (System.ArgumentException e)
            {
                Console.WriteLine("File name for overshooters is incorret");
            }
            

            //Clean Data Grapped for the overshooters file

            for (int i = 1; i < (arrayOverShooters.Length/2); i++)
            {
                if (arrayOverShooters[i][0].IndexOf("M", StringComparison.CurrentCultureIgnoreCase) != 0)
                {
                    if (!Char.IsLetter(arrayOverShooters[i][0].ElementAt(0)))
                    {
                        string str = "";
                        int len = arrayOverShooters[i][0].Length;
                        for (int m = len; m < 5; m++)
                        {
                            str = "0" + str;
                        }
                        arrayOverShooters[i][0] = "G" + str + arrayOverShooters[i][0];
                    }
                    else
                    {
                        arrayOverShooters[i][0] = "G" + arrayOverShooters[i][0];
                    }
                }
            }
            //Print Overshooting cells
            int z = 0;
            foreach (var cell in arrayOverShooters)
            {
                if (arrayCell.Contains(cell[0]))
                {
                    Console.WriteLine(cell[0] + " \tTilt:" + cell[1]);
                    z++;
                }
            }
            Console.WriteLine(z);
            */
            Console.ReadLine();

        }
    }
}
