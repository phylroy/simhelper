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
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.IO;
using DOE22; 
namespace SimHelper
{
    public partial class Form1 : Form
    {
        DOE22SimFile doe = new DOE22SimFile();
        public Form1()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Open SIM File";
            openFileDialog1.Filter = "SIM Files|*.sim";
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog and get result.
            if (result == DialogResult.OK) // Test result.
            {
                //Read in Sim file. 
                doe.scanSimFile(openFileDialog1.FileName);
                //Assign Datatables to grids. 
                dataGridView2.DataSource = doe.ZoneAnnualTable;
                dataGridView1.DataSource = doe.SystemAnnualTable;
                dataGridView3.DataSource = doe.bepsTable;
                //Load sim file in editor. 
                richTextBox1.LoadFile(openFileDialog1.FileName, RichTextBoxStreamType.PlainText);
                //Create a stack chart using beps table. 
                CreateStackedChart(doe.bepsTable, chart1);
                //output to excel
                //DataTableWriter.test(doe);

            }
        }

        private void CreateStackedChart(DataTable table, Chart chart)
        {
            chart.Series.Clear();
            foreach (DataRow row in table.Rows)
            {

                // For each Row add a new series
                string seriesName = row["meter"].ToString();
                chart.Series.Add(seriesName);
                chart.Series[seriesName].ChartType = SeriesChartType.StackedColumn;
                chart.Series[seriesName].BorderWidth = 2;

                for (int colIndex = 2; colIndex < table.Columns.Count; colIndex++)
                {
                    // For each column (column 1 and onward) add the value as a point
                    string columnName = table.Columns[colIndex].ColumnName;
                    double YVal = (double)row[columnName];
                    chart.Series[seriesName]["StackedGroupName"] = "Group1";
                    chart.Series[seriesName]["StackedGroupName"] = "Group1";
                    chart.Series[seriesName].Points.AddXY(columnName, YVal);
                }
                chart.ChartAreas["Default"].AxisX.Interval = 1;
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {

            DataGridViewColumn newColumn = dataGridView1.Columns[e.ColumnIndex];
            DataColumn ColumnY = doe.SystemAnnualTable.Columns[newColumn.HeaderText];
            DataColumn ColumnX = doe.SystemAnnualTable.Columns[0];
            DataView systemview = new DataView(doe.SystemAnnualTable);
            chart2.Series.Clear();
            chart2.Series.Add(newColumn.HeaderText);
            for (int rowIndex = 2; rowIndex < doe.SystemAnnualTable.Rows.Count; rowIndex++)
            {
                // For each column (column 1 and onward) add the value as a point


                chart2.Series[newColumn.HeaderText].Points.DataBindXY(systemview, "SYSTEM-NAME", systemview, newColumn.HeaderText);
            }
            chart2.ChartAreas["Default"].AxisX.IsLabelAutoFit = true;

        }

        private void splitContainer1_Panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
 