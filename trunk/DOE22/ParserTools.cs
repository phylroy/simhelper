using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;
using System.Text.RegularExpressions;
using System.Data;

namespace DOE22
{
    class ParserTools
    {
        public static List<string> addColumns(ref DataTable table, string[] HeaderList)
        {
            List<string> list = new List<string>();
            for (int i = 0; i < HeaderList.Count(); i = i + 2)
            {

                string test2 = HeaderList[i + 1];
                if (table.PrimaryKey.Count() == 0 | HeaderList[i + 1] != table.PrimaryKey[0].ToString())
                {
                    DataColumn column = new DataColumn(HeaderList[i + 1], Type.GetType(HeaderList[i]));
                    table.Columns.Add(column);
                    list.Add(HeaderList[i + 1]);
                }
            }
            return list;
        }

        public static void addColumn(ref DataTable table, ref List<string> list, string type, string header)
        {
            DataColumn column = new DataColumn(header, Type.GetType(type));
            table.Columns.Add(column);
        }

        public static List<string> GetRowsDataToList(string[] lineArray, int lineStartNumber, string[] regexList)
        {
            List<string> returnString = new List<string>();

            //Get Rows data from multiple lines and return a string[].
            for (int i = 0; i < regexList.Count(); i = i + 2)
            {
                int lineNumber = Convert.ToInt16(regexList[i]);
                string regexPattern = regexList[i + 1];
                List<string> tempvalues = getStringMatches(regexPattern, lineArray[lineStartNumber + lineNumber]);
                returnString.AddRange(tempvalues);
            }
            return returnString;
        }

        public static List<string> getStringMatches(string pattern, string line)
        {
            Regex regex = new Regex(pattern, RegexOptions.IgnoreCase);
            List<string> returnString = new List<string>();
            MatchCollection matches = regex.Matches(line);

            foreach (Match match in matches)
            {
                foreach (Group group in match.Groups)
                {
                    returnString.Add(group.Value.Trim());
                }
            }
            returnString.RemoveAt(0);
            return returnString;
        }



    }
}
