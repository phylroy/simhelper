using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace SimHelper
{
    class DOE22SimManager
    {
        public List<DOE22SimFile> DOESimFiles;

        //Excel Objects. 
        Excel.Application oXL;
        Excel.Workbook oWB;
        Excel.Worksheet oSheet;
        Excel.Sheets oXLSheets;
        Excel.Range oRange;


        //Flags to 
        bool bWriteESD;
        bool bWriteBEPS;
        bool bWriteZoneAnnualData;
        bool bWriteSystemAnnualData;
        string sFoldername;


        //This is the constructor for the Simulation manager. 
        public DOE22SimManager(string foldername, List<string> simFileList, bool writeBEPS, bool writeESD, bool writeZoneAnnualData, bool writeSystemAnnualData )

        {

            bWriteZoneAnnualData = writeZoneAnnualData;
            bWriteSystemAnnualData = writeSystemAnnualData;
            bWriteBEPS = writeBEPS;
            bWriteESD = writeESD;
            
            sFoldername = foldername;

            //Initialize the DOE22SIMfile. 
            DOESimFiles = new List<DOE22SimFile>();


            //Load all files from list. 
            LoadFilesFromList(simFileList);



            //Write to excel.
            WriteToExcel();



        }

        // This method loads all the simfile into the DOESimFiles Container. 
        public void LoadFilesFromFolder(string foldername)
        {

            // find all sim files in current folder. 
            DirectoryInfo di = new DirectoryInfo(foldername);
            FileInfo[] rgFiles = di.GetFiles("*.sim");
            foreach (FileInfo fi in rgFiles)
            {
                DOE22SimFile test = new DOE22SimFile(fi.FullName);
                DOESimFiles.Add(test);
                Console.Write(fi.FullName + "\n");
            }
        }

        public void LoadFilesFromList(List<string> simFileList)
        {
            foreach (string  fi in simFileList)
            {
                DOE22SimFile test = new DOE22SimFile(fi);
                DOESimFiles.Add(test);
                Console.Write(fi + "\n");
            }


        }



        //This method wirtes the sim data to the excel file. 
        public void WriteToExcel()
        {
            int iSimRunNumber = 0;
            int linenumber = 0;

            // Start Excel and get Application object. 
            oXL = new Excel.Application();

            // Set some properties 
            oXL.Visible = false;
            oXL.DisplayAlerts = false;

            // Get a new workbook. 
            oWB = oXL.Workbooks.Add(Missing.Value);

            //Add a new sheets object.
            oXLSheets = oXL.Sheets as Excel.Sheets;


            foreach (DOE22SimFile simfile in DOESimFiles)
            {
                iSimRunNumber++;
                oSheet = (Excel.Worksheet)oXLSheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                oSheet.Name = "RUN-" + iSimRunNumber.ToString();
                
                //oSheet.Name = Path.GetFileName(simfile.filepath);
                linenumber = 0;
                // Output BEPS to excel Sheet.
                oSheet.Cells[linenumber = 1, 1] = Path.GetFileName(simfile.filepath);
                linenumber++;
                oSheet.Cells[linenumber, 1] = "BEPS";
                linenumber++;
                //print bpes report. 
                PrintTableToExcel(linenumber, simfile.bepsTable, oSheet);
                linenumber = linenumber + simfile.bepsTable.Rows.Count + 1;
                linenumber++;
                oSheet.Cells[linenumber, 1] = "ES-D";
                linenumber++;
                //Print es-d report. 
                PrintTableToExcel(linenumber, simfile.esdTable, oSheet);

                // Resize the columns 
                oRange = oSheet.get_Range(oSheet.Cells[1, 1],
                                          oSheet.Cells[simfile.bepsTable.Rows.Count,
                                          simfile.bepsTable.Columns.Count]);
                oRange.EntireColumn.AutoFit();
            }

            //reset linenumber for All sheet.
            linenumber = 0;
            oSheet = (Excel.Worksheet)oXLSheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            oSheet.Name = "ALL";
            
            foreach (DOE22SimFile simfile in DOESimFiles)
            {
                linenumber++;
                // Output Filename to excel Sheet.
                oSheet.Cells[linenumber, 1] = Path.GetFileName(simfile.filepath);
                linenumber++;

                if (bWriteBEPS == true)
                {
                    // Output Filename to excel Sheet.
                    oSheet.Cells[linenumber, 1] = "BEPS";
                    linenumber++;
                    //print beps report. 
                    PrintTableToExcel(linenumber, simfile.bepsTable, oSheet);
                    linenumber = linenumber + simfile.bepsTable.Rows.Count + 1;
                }

                //Print ES-D
                if (bWriteESD == true)
                {
                    linenumber++;
                    oSheet.Cells[linenumber, 1] = "ES-D";
                    linenumber++;
                    //Print es-d report. 
                    PrintTableToExcel(linenumber, simfile.esdTable, oSheet);
                    linenumber = linenumber + simfile.esdTable.Rows.Count + 1;
                }

                //Print Zone Annual Data
                if (bWriteZoneAnnualData == true)
                {
                    linenumber++;
                    oSheet.Cells[linenumber, 1] = "Zone Annual Data";
                    linenumber++;
                    //Print Zone Annual Data report. 
                    PrintTableToExcel(linenumber, simfile.ZoneAnnualTable, oSheet);
                    linenumber = linenumber + simfile.ZoneAnnualTable.Rows.Count + 1;
                }

                //Print System Annual Data
                if (bWriteSystemAnnualData == true)
                {
                    linenumber++;
                    oSheet.Cells[linenumber, 1] = "System Annual Data";
                    linenumber++;
                    //Print Zone Annual Data report. 
                    PrintTableToExcel(linenumber, simfile.SystemAnnualTable, oSheet);
                    linenumber = linenumber + simfile.SystemAnnualTable.Rows.Count + 1;
                }




                // Resize the columns 
                oRange = oSheet.get_Range(oSheet.Cells[1, 1],
                                          oSheet.Cells[simfile.bepsTable.Rows.Count,
                                          simfile.bepsTable.Columns.Count]);
                oRange.EntireColumn.AutoFit();
            }
            // Save the sheet and close 
            oSheet = null;
            oRange = null;
            oWB.SaveAs(sFoldername + @"\test.xls", Excel.XlFileFormat.xlWorkbookNormal,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Excel.XlSaveAsAccessMode.xlExclusive,
                Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value);
            oWB.Close(Missing.Value, Missing.Value, Missing.Value);
            oWB = null;

            // Clean up 
            // NOTE: When in release mode, this does the trick 
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
        private static void PrintTableToExcel(int RowNumber, DataTable dt, Excel.Worksheet oSheet)
        {
            int rowCount = 0;
            foreach (DataRow dr in dt.Rows)
            {
                rowCount += 1;
                for (int i = 1; i < dt.Columns.Count + 1; i++)
                {
                    // Add the header the first time through 
                    if (rowCount == 2)
                    {
                        oSheet.Cells[RowNumber, i] = dt.Columns[i - 1].ColumnName;
                    }
                    oSheet.Cells[rowCount + RowNumber, i] = dr[i - 1].ToString();
                }
            }
        }

    }
}
