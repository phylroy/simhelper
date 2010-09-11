//  Copyright (C) 2008-2010 Phylroy Lopez

// This program is free software; you can redistribute it and/or modify it
// under the terms of the GNU General Public License as published by the
// Free Software Foundation; either version 2, or (at your option) any
// later version.

// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.

// http://www.gnu.org/licenses/gpl.html

using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Text;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel; 
using System.Reflection;
using DOE22;

namespace DOE22
{
    class DataTableWriter
    {
        public static void test(DOE22SimFile doe2simdata)
        {
        Excel.Application oXL; 
Excel.Workbook oWB; 
Excel.Worksheet oSheet; 
Excel.Range oRange;

// Start Excel and get Application object. 
oXL = new Excel.Application();

// Set some properties 
oXL.Visible = true; 
oXL.DisplayAlerts = false;

// Get a new workbook. 
oWB = oXL.Workbooks.Add(Missing.Value);

// Get the active sheet 
oSheet = (Excel.Worksheet)oWB.ActiveSheet ;
oSheet.Name = Path.GetFileName(doe2simdata.filepath); 

// Process the DataTable 
// BE SURE TO CHANGE THIS LINE TO USE *YOUR* DATATABLE 
DataTable dt = doe2simdata.bepsTable;

int rowCount = 1; 
foreach (DataRow dr in dt.Rows) 
{ 
    rowCount += 1; 
    for (int i = 1; i < dt.Columns.Count+1; i++) 
    { 
        // Add the header the first time through 
        if (rowCount==2) 
        { 
            oSheet.Cells[1, i] = dt.Columns[i - 1].ColumnName; 
        } 
        oSheet.Cells[rowCount, i] = dr[i - 1].ToString(); 
    } 
}

DataTable SystemTable = doe2simdata.SystemAnnualTable;

rowCount =  1;
foreach (DataRow dr in SystemTable.Rows)
{
    rowCount += 1;
    for (int i = 1; i < SystemTable.Columns.Count + 1; i++)
    {
        // Add the header the first time through 
        if (rowCount == 2)
        {
            oSheet.Cells[SystemTable.Rows.Count, i] = SystemTable.Columns[i - 1].ColumnName;
        }
        oSheet.Cells[rowCount + SystemTable.Rows.Count, i] = dr[i - 1].ToString();
    }
}




// Resize the columns 
oRange = oSheet.get_Range(oSheet.Cells[1, 1], 
              oSheet.Cells[rowCount, dt.Columns.Count]); 
oRange.EntireColumn.AutoFit();
// Create Chart
Excel.Range chartRange;

Excel.ChartObjects xlCharts = (Excel.ChartObjects)oSheet.ChartObjects(Type.Missing);
Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 300, 250);
Excel.Chart chartPage = myChart.Chart;

Excel.SeriesCollection seriesCollection = (Excel.SeriesCollection)chartPage.SeriesCollection(Type.Missing);
Excel.Series series1 = seriesCollection.NewSeries();
Excel.Series series2 = seriesCollection.NewSeries();
Excel.Series series3 = seriesCollection.NewSeries();
series1.Values = oSheet.get_Range("C1", "N1");
series2.Values = oSheet.get_Range("C2", "N2");



//chartPage.SetSourceData(seriesCollection, Excel.XlRowCol.xlColumns);
chartPage.ChartType = Excel.XlChartType.xlColumnClustered;



// Save the sheet and close 
oSheet = null; 
oRange = null; 
oWB.SaveAs("test.xls", Excel.XlFileFormat.xlWorkbookNormal, 
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

    }
}
