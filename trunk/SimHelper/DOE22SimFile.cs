using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;
using System.Text.RegularExpressions;
using System.Data;
namespace SimHelper
{
    class DOE22SimFile
    {
        //Members.
        //location of file path. 
        public string filepath;
        public DataTable SystemAnnualTable;
        public DataTable ZoneAnnualTable;

        //BEPS table and header list vars.
        public DataTable bepsTable;
        private List<string> bepsHeaderList;

        //ES-D table and header list vars.
        public DataTable esdTable;
        private List<string> esdHeaderList;

        //Zone table for Annual or static data. 
        //public DataTable ZoneAnnualTable;

        //SV-A Header lists. 
        private List<string> zoneSVAHeaderList;
        private List<string> systemSVAHeaderList1;
        private List<string> systemSVAHeaderList2;
        //SS-R Header lists. 
        private List<string> zoneSSRHeaderList;



        private void scanSimFile()
        {

            //clear all tables. 
            Initialize();
            string[] lineArray = File.ReadAllLines(filepath);
            for (int linecounter = 0; linecounter < lineArray.Count(); linecounter++)
            {

                getBEPS(lineArray, linecounter);
                getESD(lineArray, linecounter);
                getSSR(lineArray, linecounter);
            }

            //Because the SV-A report truncates to the 25 char, the zone names must be initilized before the SV-A report.
            for (int linecounter = 0; linecounter < lineArray.Count(); linecounter++)
            {
                string line = lineArray[linecounter];
                getSVA(lineArray, linecounter);
            }
        }

        public void scanSimFile(string filepath)
        {
            this.filepath = filepath;
            scanSimFile();
        }

        public DOE22SimFile()
        {
            Initialize();

        }

        public DOE22SimFile(string filepath)
        {
            Initialize();
            scanSimFile(filepath);

        }



        private void Initialize()
        {

            bepsTable = new DataTable("bepsTable");
            bepsTable.Columns.Add("METER", typeof(string));
            bepsTable.Columns["METER"].Unique = true;
            bepsTable.PrimaryKey = new DataColumn[] { bepsTable.Columns["METER"] };
            bepsHeaderList = ParserTools.addColumns(ref bepsTable, new string[]{
                "System.String","METER",
                "System.String","FUEL",
                "System.Double","LIGHTS",
                "System.Double","TASK-LIGHTS",
                "System.Double","MISC-EQUIP",
                "System.Double","SPACE-HEATING",
                "System.Double","SPACE-COOLING",
                "System.Double","HEAT-REJECTION",
                "System.Double","PUMPS-AUX",
                "System.Double","VENT-FANS",
                "System.Double","REFRIG-DISPLAY",
                "System.Double","HT-PUMP-SUPPLEM",
                "System.Double","DOMEST-HOT-WTR",
                "System.Double","EXT-USAGE",
                "System.Double","TOTAL"
            });

            esdTable = new DataTable("esdTable");
            esdTable.Columns.Add("UTILITY-RATE", typeof(string));
            esdTable.Columns["UTILITY-RATE"].Unique = true;
            esdTable.PrimaryKey = new DataColumn[] { esdTable.Columns["UTILITY-RATE"] };
            esdHeaderList = ParserTools.addColumns(ref esdTable, new string[]{
                "System.String","UTILITY-RATE",
                "System.String","RESOURCE",
                "System.String","METERS",
                "System.Double","METERED-ENERGY UNITS/YR",
                "System.String","UNIT",
                "System.Double","TOTAL CHARGE ($)",
                "System.Double","VIRTUAL RATE ($/UNIT)",
                "System.String","RATE USED ALL YEAR"
            });



            SystemAnnualTable = new DataTable("systemTable");
            SystemAnnualTable.Columns.Add("SYSTEM-NAME", typeof(string));
            SystemAnnualTable.Columns["SYSTEM-NAME"].Unique = true;
            SystemAnnualTable.PrimaryKey = new DataColumn[] { SystemAnnualTable.Columns["SYSTEM-NAME"] };
            //First line.
            systemSVAHeaderList1 = ParserTools.addColumns(ref SystemAnnualTable, new string[] {    
                "System.String", "SYSTEM-NAME",
                "System.String", "SYSTEM-TYPE",
                "System.Double", "ALTITUDE-FACTOR",
                "System.Double", "FLOOR-AREA",
                "System.Double", "MAX-PEOPLE",
                "System.Double", "OUTSIDE-AIR-RATIO",
                "System.Double", "COOLING-CAPACITY",
                "System.Double", "SENSIBLE",
                "System.Double", "HEATING-CAPACITY",
                "System.Double", "COOLING-EIR",
                "System.Double", "HEATING-EIR",
                "System.Double", "HEAT PUMP-SUPP-HEAT"
            });

            //Second line
            systemSVAHeaderList2 = ParserTools.addColumns(ref SystemAnnualTable, new string[]{
                "System.String", "FAN-TYPE",
                "System.Double", "FAN-CAPACITY",
                "System.Double", "DIVERSITY-FACTOR",
                "System.Double", "POWER-DEMAND",
                "System.Double", "FAN-DELTA-T",
                "System.Double", "STATIC-PRESSURE",
                "System.Double", "TOTAL-EFF",
                "System.Double", "MECH-EFF",
                "System.String", "FAN-PLACEMENT",
                "System.String", "FAN_CONTROL",
                "System.Double", "MAX-FAN-RATIO",
                "System.Double", "MIN-FAN-RATIO"
            });

            ZoneAnnualTable = new DataTable("zoneTable");
            ZoneAnnualTable.Columns.Add("ZONE-NAME", typeof(string));
            ZoneAnnualTable.Columns["ZONE-NAME"].Unique = true;
            ZoneAnnualTable.PrimaryKey = new DataColumn[] { ZoneAnnualTable.Columns["ZONE-NAME"] };
            zoneSVAHeaderList = ParserTools.addColumns(ref ZoneAnnualTable, new string[]{
                "System.String", "ZONE-NAME",
                "System.Double", "SUPPLY-FLOW",
                "System.Double", "EXHAUST-FLOW",
                "System.Double", "FAN",
                "System.Double", "MINIMUM-FLOW",
                "System.Double", "OUTSIDE-AIR-FLOW",
                "System.Double", "COOLING-CAPACITY",
                "System.Double", "SENSIBLE",
                "System.Double", "EXTRACTION-RATE",
                "System.Double", "HEATING-CAPACITY",
                "System.Double", "ADDITION-RATE",
                "System.Double", "ZONE-MULT",
                "System.Double", "BASEBOARD-HEATING-CAPACITY"
            });


            //SS-R report.


            //            VECC 2008 Proposed Model DDMode                                                  DOE-2.2-44e4   2/23/2010    16:29:18  BDL RUN  1

            //REPORT- SS-R Zone Performance Summary for   RTU-1 (PSZ)                                     WEATHER FILE- VANCOUVER TMY       
            //---------------------------------------------------------------------------------------------------------------------------------


            //                   ZONE OF  ZONE OF   ZONE     ZONE        --------  Number of hours within each PART LOAD range  --------- TOTAL
            //                   MAXIMUM  MAXIMUM   UNDER    UNDER         00    10    20    30    40    50    60    70    80    90   100   RUN
            //                  HTG DMND CLG DMND  HEATED   COOLED         10    20    30    40    50    60    70    80    90   100    +  HOURS
            // ZONE              (HOURS)  (HOURS)  (HOURS)  (HOURS)
            //----------------  -------- -------- -------- --------      ----  ----  ----  ----  ----  ----  ----  ----  ----  ----  ----  ----
            //EL4 South Perim Zn (G.S11)      
            //                         0        0        0        0         0     0     0  1226     0     0     0     0     0     0  3840  5066
            //EL1 South Perim Zn (G.S11)      
            //                         0        0        0      623         0     0     0  1152   374   436   425   639   353  1687     0  5066

            //                  -------- -------- -------- --------

            //           TOTAL         0        0        0      623

            zoneSSRHeaderList = ParserTools.addColumns(ref ZoneAnnualTable, new string[]{
                "System.String", "ZONE-NAME",
                "System.Double", "ZONE-OF-MAX-HTG-DMND (HOURS)",
                "System.Double", "ZONE-OF-MAX-CLG-DMND (HOURS)",
                "System.Double", "ZONE-UNDER-HEATED (HOURS)",
                "System.Double", "ZONE-UNDER-COOLED (HOURS)",
                "System.Double", "NUM-HOURS PLR 00-10",
                "System.Double", "NUM-HOURS PLR 10-20",
                "System.Double", "NUM-HOURS PLR 20-30",
                "System.Double", "NUM-HOURS PLR 30-40",
                "System.Double", "NUM-HOURS PLR 40-50",
                "System.Double", "NUM-HOURS PLR 50-60",
                "System.Double", "NUM-HOURS PLR 60-70",
                "System.Double", "NUM-HOURS PLR 70-80",
                "System.Double", "NUM-HOURS PLR 80-90",
                "System.Double", "NUM-HOURS PLR 90-100",
                "System.Double", "NUM-HOURS PLR 100-+",
                "System.Double", "NUM-HOURS PLR Total"
        });
        }

        //method stores the beps energy for the run. 
        private void getBEPS(string[] lineArray, int linecounter)
        {

            //Search for BEPS report.


            //not using regex here..it is too slow. 
            if (lineArray[linecounter].IndexOf("REPORT- BEPS Building Energy Performance") >= 0)
            {
                //find last line of BEPS Energy table. 
                int lastline = linecounter;
                while (lineArray[lastline] != "              =======  =======  =======  =======  =======  =======  =======  =======  =======  =======  =======  =======  ========")
                {
                    lastline += 1;
                }

                //Start reading in BEPS energy, eight lines from start of beps.
                int sublinecounter = linecounter + 7;
                while (sublinecounter < lastline)
                {

                    //get fuel meter name
                    string[] regexList = 
                        {  
                            "0",@"^(\w*)\s+(\S*)\s*$",
                            "1",@".{12}(.{9})(.{9})(.{9})(.{9})(.{9})(.{9})(.{9})(.{9})(.{9})(.{9})(.{9})(.{9})?(.*)?"
                        };
                    string linetest1 = lineArray[sublinecounter];
                    string linetest2 = lineArray[sublinecounter + 1];
                    string linetest3 = lineArray[sublinecounter + 2];
                    string linetest4 = lineArray[sublinecounter + 3];
                    List<string> values = ParserTools.GetRowsDataToList(lineArray, sublinecounter, regexList);
                    sublinecounter += 3;
                    DataRow row;
                    DataTable tabletest = bepsTable;
                    string valuetest = values[0];
                    if (tabletest.Rows.Find(values[0]) == null)
                    {
                        row = tabletest.NewRow();
                        row["METER"] = values[0];
                        values.RemoveAt(0);
                    }
                    else
                    {
                        row = tabletest.Rows.Find(values[0]);
                        values.RemoveAt(0);
                    }


                    for (int iCounter = 0; iCounter < values.Count(); iCounter++)
                    {
                        row[bepsHeaderList[iCounter]] = values[iCounter];
                    }
                    bepsTable.Rows.Add(row);
                }
                bepsTable.AcceptChanges();

            }
        }

        //method stores the esd energy for the run. 
        private void getESD(string[] lineArray, int linecounter)
        {
            string line = lineArray[linecounter];
            //Search for esd report.
            if (lineArray[linecounter].IndexOf("REPORT- ES-D Energy Cost Summary") >= 0)
            {
                //find last line of esd Energy table. 
                int lastline = linecounter;
                while (lineArray[lastline] != "                                                                                          ==========")
                {
                    lastline += 1;
                }
                //Start reading in esd energy, six lines from start of esd.
                int sublinecounter = linecounter + 8;
                while (sublinecounter < lastline)
                {
                    //get fuel meter name
                    string[] regexList = 
                        {  
                            "0",@"(.{35})(.{19})(.{11})(.{13})(.{12})(.{10})(.{13})(.{9})"
                        };
                    List<string> values = ParserTools.GetRowsDataToList(lineArray, sublinecounter, regexList);
                    sublinecounter += 2;
                    DataRow row;
                    DataTable tabletest = esdTable;
                    string valuetest = values[0];
                    //Check to see if a row does not exist..if not create a new row.
                    if (tabletest.Rows.Find(values[0]) == null)
                    {
                        row = tabletest.NewRow();
                        row["UTILITY-RATE"] = values[0];
                        values.RemoveAt(0);
                    }
                    else
                    {
                        //otherwise find the existing row and use it. 
                        row = tabletest.Rows.Find(values[0]);
                        values.RemoveAt(0);
                    }
                    for (int iCounter = 0; iCounter < values.Count(); iCounter++)
                    {
                        row[esdHeaderList[iCounter]] = values[iCounter];
                    }
                    esdTable.Rows.Add(row);
                }
                esdTable.AcceptChanges();
            }
        }

        //method stores the beps energy for the run. 
        private void getSVA(string[] lineArray, int linecounter)
        {
            

            //Search for SV-A
            if (lineArray[linecounter].IndexOf("REPORT- SV-A System Design Parameters for") >= 0)
            {

                Regex sva_flag = new Regex(@"^\s*REPORT- SV-A System Design Parameters for(.{51})WEATHER FILE.*$", RegexOptions.IgnoreCase);
                if (sva_flag.IsMatch(lineArray[linecounter]))
                {
                    //Zone 
                    DataRow row;
                    MatchCollection matches = sva_flag.Matches(lineArray[linecounter]);
                    string system_name = matches[0].Groups[1].Value.ToString().Trim();

                    //Start reading in System data, six lines from start of beps.
                    int sublinecounter = linecounter + 6;
                    string line1 = lineArray[sublinecounter];

                    List<string> values = ParserTools.getStringMatches(@"\s*(.{8})(.{11})(.{11})(.{11})(.{11})(.{11})(.{11})(.{11})(.{11})(.{11})(.{11})\s*$", lineArray[sublinecounter]);
                    values.Insert(0, system_name);
                    if (SystemAnnualTable.Rows.Find(values[0]) == null)
                    {
                        row = SystemAnnualTable.NewRow();
                        row["SYSTEM-NAME"] = values[0];
                        values.RemoveAt(0);
                        SystemAnnualTable.Rows.Add(row);
                    }
                    else
                    {
                        row = SystemAnnualTable.Rows.Find(values[0]);
                        values.RemoveAt(0);
                    }

                    int counter = 0;
                    foreach (string value in values)
                    {
                        row[systemSVAHeaderList1[counter]] = value;
                        counter += 1;
                    }
                    int iNotSumLine = 0;
                    if (values[0] != "SUM")
                    {

                        //Start reading in System data line 2.
                        sublinecounter = linecounter + 13;
                        line1 = lineArray[sublinecounter];
                        List<string> values2 = ParserTools.getStringMatches(@"\s*(.{8})(.{11})(.{11})(.{9})(.{10})(.{11})(.{8})(.{8})(.{12})(.{10})(.{10})(.{10})\s*$", lineArray[sublinecounter]);
                        int counter2 = 0;
                        foreach (string value in values2)
                        {
                            row[systemSVAHeaderList2[counter2]] = value;
                            counter2 += 1;
                        }
                        iNotSumLine = 4;
                    }

                    SystemAnnualTable.AcceptChanges();


                    //Start reading in Zone data.
                    sublinecounter = linecounter + 16 + iNotSumLine;

                    Regex end_flag = new Regex(@".*DOE-.*\d+/\d+/\d+.*BDL RUN.*$", RegexOptions.IgnoreCase);
                    Regex baseboard = new Regex(@"\s*(.*)\(BASEBOARDS\)$", RegexOptions.IgnoreCase);
                    while (!end_flag.IsMatch(lineArray[sublinecounter]))
                    {
                        int skip = 0;
                        if (lineArray[sublinecounter] != "")
                        {


                            line1 = lineArray[sublinecounter];
                            List<string> values3 = ParserTools.getStringMatches(@"(.{26})(.{10})(.{10})(.{10})(.{10})(.{10})(.{10})(.{10})(.{10})(.{10})(.{10})(.{5})$", lineArray[sublinecounter]);
                            if (baseboard.IsMatch(lineArray[sublinecounter + 1]))
                            {
                                List<string> baseboardVal = ParserTools.getStringMatches(@"\s*(.*)\(BASEBOARDS\)$", lineArray[sublinecounter + 1]);
                                values3.Add(baseboardVal[0]);
                                skip = 1;
                            }
                            else
                            {
                                values3.Add("0.0");
                                skip = 0;
                            }

                            row = null;
                            for (int i = 0; i < ZoneAnnualTable.Rows.Count; i++)
                            {
                                DataRow testrow = ZoneAnnualTable.Rows[i];
                                string Zonename = testrow["ZONE-NAME"].ToString();

                                if (Zonename == values3[0]
                                    || (Zonename.Length >= 25 && Zonename.Substring(0, 25) == values3[0]))
                                {
                                    row = testrow;
                                }
                            }

                            if (row == null)
                            {
                                row = ZoneAnnualTable.NewRow();
                                row["ZONE-NAME"] = values3[0];
                                ZoneAnnualTable.Rows.Add(row);
                            }
                            values3.RemoveAt(0);
                            if (values3.Count == 12)
                            {
                                int counter3 = 0;
                                foreach (string value in values3)
                                {
                                    row[zoneSVAHeaderList[counter3]] = value;
                                    counter3 += 1;
                                }
                                ZoneAnnualTable.AcceptChanges();
                            }
                        }
                        sublinecounter += (1 + skip);
                    }
                }
            }
        }
        //method stores the beps energy for the run. 
        private void getSSR(string[] lineArray, int linecounter)
        {

            //Search for SV-A
            if (lineArray[linecounter].IndexOf("REPORT- SS-R Zone Performance") >= 0)
            {

                string line = lineArray[linecounter];

                //Search for BESS-R report.
                string pattern = @"(\s*REPORT- SS-R Zone Performance.*$)";
                Regex ssr_flag = new Regex(pattern, RegexOptions.IgnoreCase);
                if (ssr_flag.IsMatch(line))
                {
                    //find last line of BEPS Energy table. 
                    int lastline = linecounter;
                    while (lineArray[lastline] != "                  -------- -------- -------- --------")
                    {
                        lastline += 1;
                    }

                    //Start reading in BEPS energy, nine lines from start of beps.
                    int sublinecounter = linecounter + 9;
                    while (sublinecounter + 2 < lastline)
                    {

                        //get fuel meter name
                        string[] regexList = 
                        {  
                            "0",@"^(.*)$",
                            "1",@"^(.{26})(.{9})(.{9})(.{9})(.{10})(.{6})(.{6})(.{6})(.{6})(.{6})(.{6})(.{6})(.{6})(.{6})(.{6})(.{6})$"
                        };

                        List<string> values = ParserTools.GetRowsDataToList(lineArray, sublinecounter, regexList);
                        sublinecounter += 2;
                        DataRow row;
                        DataRow foundrow = null;
                        DataTable tabletest = ZoneAnnualTable;
                        string valuetest = values[0];

                        for (int i = 0; i < ZoneAnnualTable.Rows.Count; i++)
                        {
                            DataRow testrow = ZoneAnnualTable.Rows[i];
                            string Zonename = testrow["ZONE-NAME"].ToString();

                            if (Zonename == values[0] || (Zonename.Length == 25 && Zonename.Substring(0, 25) == values[0]))
                            {
                                foundrow = testrow;
                            }
                        }
                        if (foundrow == null)
                        {
                            row = ZoneAnnualTable.NewRow();
                            row["ZONE-NAME"] = values[0];
                            values.RemoveAt(0);
                            ZoneAnnualTable.Rows.Add(row);

                        }
                        else
                        {
                            row = foundrow;
                            row = ZoneAnnualTable.Rows.Find(values[0]);
                            values.RemoveAt(0);
                        }


                        for (int iCounter = 0; iCounter < values.Count(); iCounter++)
                        {
                            row[zoneSSRHeaderList[iCounter]] = values[iCounter];
                        }
                        //zoneTable.Rows.Add(row);
                    }
                    ZoneAnnualTable.AcceptChanges();
                }
            }
        }
    }
}


