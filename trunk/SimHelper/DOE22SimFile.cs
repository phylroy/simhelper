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


        public DataTable BuildingAnnualTable;
        public DataTable BuildingMonthlyTable;
        public DataTable BuildingHourlyTable;


        public DataTable PlantAnnualTable;
        public DataTable[] PlantMonthlyTables;
        public DataTable[] PlantHourlyTables;

        public DataTable SystemAnnualTable;
        public DataTable[] SystemMonthlyTables;
        public DataTable[] SystemHourlyTables;

        public DataTable ZoneAnnualTable;
        public DataTable[] ZoneMonthlyTables;
        public DataTable[] ZoneHourlyTables;

        public DataTable SpaceAnnualTable;
        public DataTable[] SpaceMonthlyTables;
        public DataTable[] SpaceHourlyTables;



        //BEPS table and header list vars.
        public DataTable bepsTable;
        public List<string> bepsHeaderList;

        //Zone table for Annual or static data. 
        //public DataTable ZoneAnnualTable;



        //SV-A Header lists. 
        public List<string> zoneSVAHeaderList;
        public List<string> systemSVAHeaderList1;
        public List<string> systemSVAHeaderList2;
        //SS-R Header lists. 
        public List<string> zoneSSRHeaderList;

        public void scanSimFile()
        {
            //clear all tables. 
            Initialize();
            string[] lineArray = File.ReadAllLines(filepath);
            for (int linecounter = 0; linecounter < lineArray.Count(); linecounter++)
            {
                string line = lineArray[linecounter];
                getBEPS(lineArray, linecounter);

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

        //Constructor.
        public void Initialize()
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
        public void setFilePath(string inpfilepath)
        {
            this.filepath = inpfilepath;
        }

        //method stores the beps energy for the run. 
        public void getBEPS(string[] lineArray, int linecounter)
        {
            string line = lineArray[linecounter];

            //Search for BEPS report.
            string pattern = @"(\s*REPORT- BEPS Building Energy Performance.*$)";
            Regex beps_flag = new Regex(pattern, RegexOptions.IgnoreCase);
            if (beps_flag.IsMatch(line))
            {
                //find last line of BEPS Energy table. 
                int lastline = linecounter;
                while (lineArray[lastline] != "              =======  =======  =======  =======  =======  =======  =======  =======  =======  =======  =======  =======  ========")
                {
                    lastline += 1;
                }

                //Start reading in BEPS energy, six lines from start of beps.
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

        //method stores the beps energy for the run. 
        public void getSVA(string[] lineArray, int linecounter)
        {
            string line = lineArray[linecounter];

            //Search for SV-A
            string pattern = @"^\s*REPORT- SV-A System Design Parameters for(.{51})WEATHER FILE.*$";
            Regex sva_flag = new Regex(pattern, RegexOptions.IgnoreCase);
            if (sva_flag.IsMatch(line))
            {
                //Zone 
                DataRow row;
                MatchCollection matches = sva_flag.Matches(line);
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

                        DataRow zone_row;
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


                        ////



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

        //method stores the beps energy for the run. 
        public void getSSR(string[] lineArray, int linecounter)
        {
            string line = lineArray[linecounter];

            //Search for BEPS report.
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


