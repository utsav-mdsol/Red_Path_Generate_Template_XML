using Data = Google.Apis.Sheets.v4.Data;
using System.Threading;
using System.IO;
using Google.Apis.Util.Store;
using System.Xml.Linq;
using System.Xml;
using System.Text.RegularExpressions;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Http;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Newtonsoft.Json;
using System.Net.Http;


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OneClickCore
{
    public class Result_List
    {

        public string scriptname { get; set; }
        public string transactionname { get; set; }
        public string min { get; set; }
        public string avg { get; set; }
        public string max { get; set; }
        public string std { get; set; }
        public string count { get; set; }
        public string bound1 { get; set; }
        public string bound2 { get; set; }


        static void Main(string[] args)
        {
            getTagName();
        }



        public class Process_Results
        {


            //Get the list from here make a call to this method..
            public static List<string> getSPMergeFileCSVFormattedResult(string v_Input_XML_File_Path, out List<Result_List> Result, string v_Output_CSV_File_Path = null)
            {
                List<Result_List> Result_Output = new List<Result_List>();
                List<string> _out = new List<string>();
                try
                {
                    var content = File.ReadAllText(v_Input_XML_File_Path);
                    XDocument x = XDocument.Parse(content.ToString());
                    Result_List Result_Set = new Result_List();
                    foreach (XElement group in x.Root.Element("UserGroups").Elements("Group"))
                    {
                        string v_Script_Name = group.Element("Name").Value.ToString().Substring(0, (group.Element("Name").Value.ToString().IndexOf('/') + 0)).Trim();
                        v_Script_Name = v_Script_Name.Substring(0, v_Script_Name.LastIndexOf("."));
                        foreach (XElement Measure in group.Element("Measures").Elements("Measure"))
                        {
                            if (Measure.Element("Class").Value.ToString() == "Timer")
                            {
                                if (Measure.Element("Type").Value.ToString() == "Response time[s]")
                                {
                                    Result_Set = new Result_List();
                                    Result_Set.scriptname = v_Script_Name;
                                    Result_Set.transactionname = Measure.Element("Name").Value.ToString().Trim();
                                    Result_Set.min = Measure.Element("MinMin").Value.ToString().Substring(0, (Measure.Element("MinMin").Value.ToString().IndexOf('.') + 4)).Trim();
                                    Result_Set.avg = Measure.Element("Avg").Value.ToString().Substring(0, (Measure.Element("Avg").Value.ToString().IndexOf('.') + 4)).Trim();
                                    Result_Set.max = Measure.Element("MaxMax").Value.ToString().Substring(0, (Measure.Element("MaxMax").Value.ToString().IndexOf('.') + 4)).Trim();
                                    Result_Set.std = Measure.Element("Stdd").Value.ToString().Substring(0, (Measure.Element("Stdd").Value.ToString().IndexOf('.') + 4)).Trim();
                                    Result_Set.count = Measure.Element("SumCount1").Value.ToString().Substring(0, (Measure.Element("SumCount1").Value.ToString().IndexOf('.') + 0)).Trim();
                                    Result_Set.bound1 = Measure.Element("SumTimeboundCount1").Element("Percent").Value.ToString().Substring(0, (Measure.Element("SumTimeboundCount1").Element("Percent").Value.ToString().IndexOf('.') + 4)).Trim();
                                    Result_Set.bound2 = Measure.Element("SumTimeboundCount2").Element("Percent").Value.ToString().Substring(0, (Measure.Element("SumTimeboundCount2").Element("Percent").Value.ToString().IndexOf('.') + 4)).Trim();
                                    Result_Output.Add(Result_Set);
                                    Result_Set = null;
                                }
                            }
                        }
                    }
                    _out.Add("Script Name,Transaction Name,Minimum,Average,Maximum,Standard Deviation,Count,Bound1,Bound2");
                    foreach (Result_List n in Result_Output)
                    {
                        _out.Add(string.Format("{0},{1},{2},{3},{4},{5},{6},{7}", n.scriptname, n.transactionname, n.min, n.avg, n.max, n.std, n.count, n.bound1, n.bound2));
                    }
                    //File.WriteAllLines(@v_Output_CSV_File_Path, Result_Output.Select(n => n.scriptname.ToString() + ","
                    //+ n.transactionname.ToString() + ","
                    //+ n.min.ToString() + ","
                    //+ n.avg.ToString() + ","
                    //+ n.max.ToString() + ","
                    //+ n.std.ToString() + ","
                    //+ n.count.ToString() + ","
                    //+ n.bound1.ToString() + ","
                    //+ n.bound2.ToString()));
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error encountered : " + e.Message);
                }
                Result = new List<Result_List>(Result_Output);
                return _out;
            }
        }






        //Collects the OverallWorkflowMetrics from the XML file
        public static string getOverallWorkflowMetricsData(string nameString, string typeString, string captionString, string mergedXMLPath = "C:\\Users\\udhungel\\Downloads\\m@hdcperfagtv004@Platform_Redpath_OVR.xml")
        {
            //Gets the link to the RedPath Result XML File
            //var silkPerformerURL = silkPerformerLinkBox.Text;

            var xmlFileURL = mergedXMLPath;

            //Using XML Reader for scanning each line of code and getiing the required result
            using (var redPathXMLFile = XmlReader.Create(mergedXMLPath))
            {
                var captionValue = "";
                var reqValueTag = "";
                var reqNumberTag = "";
                var reqNumberValue = "";


                //Loops thorugh until the EOF
                while (redPathXMLFile.Read())
                {
                    //Gets the name of each element tag.
                    var IDElementTag = redPathXMLFile.LocalName;

                    //Searches for the tag <Name>
                    if (IDElementTag.Equals("Name"))
                    {

                        //Gets the inner text of the <Name> tag
                        redPathXMLFile.Read();
                        var IDInnerValue = redPathXMLFile.Value;


                        string ignoresSpaceIDInnerValue = Regex.Replace(IDInnerValue, @"\s+", String.Empty);// Removes the whitespaces, \n from the <name> inner text.

                        //Exxecutes if the pattern matches the argument nameString.
                        if (ignoresSpaceIDInnerValue.Equals(nameString))
                        {

                            //Loops through the file until it matches the <type> tag and the matching inner text of <type> tag with the argument provided.
                            do
                            {

                                captionValue = redPathXMLFile.LocalName;
                                redPathXMLFile.Read();
                                reqValueTag = redPathXMLFile.Value;

                            } while (!(captionValue.Equals("Type") && reqValueTag.Equals(typeString)));


                            //Executes if the pattern is matched. 
                            if ((captionValue.Equals("Type") && reqValueTag.Equals(typeString)))
                            {

                                // Goes to the tag whose inner text is returned.
                                do
                                {
                                    redPathXMLFile.Read();
                                    reqNumberTag = redPathXMLFile.LocalName;

                                } while (!(reqNumberTag.Equals(captionString)));


                                // IF the matching tag is found it returns the inner text.
                                if (reqNumberTag.Equals(captionString))
                                {
                                    redPathXMLFile.Read();
                                    reqNumberValue = redPathXMLFile.Value;
                                    reqNumberValue = Regex.Replace(reqNumberValue, @"\s+", String.Empty);// Removes the whitespaces, \n from the <name> inner text.
                                    if (reqNumberValue.Equals("")) //Checks to see if there is any other child nodes inside
                                    {
                                        while (reqNumberValue.Equals(""))
                                        { //if there is child node 
                                            redPathXMLFile.Read();
                                            reqNumberValue = redPathXMLFile.Value; //gets the inner value of the child nodes.
                                        };

                                    }
                                    return reqNumberValue;


                                }


                                return "N/A";
                            }




                        }


                    }



                }
                return "N/A";

            }
            // return null;

        }
        public static string getOverallTestMetricsData(string groupID, string captionString, string mergedXMLPath = "C:\\Users\\udhungel\\Downloads\\m@hdcperfagtv004@Platform_Redpath_OVR.xml")
        {
            //Gets the link to the RedPath Result XML File
            //var silkPerformerURL = silkPerformerLinkBox.Text;

            var xmlFileURL = mergedXMLPath;

            //Using XML Reader for scanning each line of code and getiing the required result
            using (var redPathXMLFile = XmlReader.Create(mergedXMLPath))
            {
                var captionValue = "";
                var reqNumberValue = "";
                //Loops thorugh until the EOF
                while (redPathXMLFile.Read())
                {

                    var IDElementTag = redPathXMLFile.LocalName;  //Gets the element tag name.

                    //Checks if the Current Node is <ID>
                    if (IDElementTag.Equals("ID"))
                    {

                        redPathXMLFile.Read();
                        var IDInnerValue = redPathXMLFile.Value;

                        //Checks to see if the inner text of ID matches the pattern provided in the parameters.
                        if (IDInnerValue.Equals(groupID))
                        {

                            //Loops through all the <ID> tag until the <caption> inner text matches the argument provided as a parameter.
                            do
                            {
                                redPathXMLFile.Read();
                                captionValue = redPathXMLFile.Value;
                            } while (!(captionValue.Equals(captionString)));//Stops when the specific caption tag is found

                            //Function to get the value after the <caption> tag.
                            if (captionValue.Equals(captionString))
                            {
                                redPathXMLFile.ReadToNextSibling("Value"); //Skips to the next node
                                redPathXMLFile.Read();
                                redPathXMLFile.Read();
                                redPathXMLFile.Read();
                                reqNumberValue = redPathXMLFile.Value; // Gets the required data
                                return reqNumberValue;

                            }

                        }

                    }

                }
                return null; //Returns null if the data not found

            }

            //return null;
        }
        public static string getRedPathSetupData(string searchTag, string mergedXMLPath = "C:\\Users\\udhungel\\Downloads\\m@hdcperfagtv004@Platform_Redpath_OVR.xml")
        {
            var xmlFileURL = mergedXMLPath;

            //Using XML Reader for scanning each line of code and getiing the required result
            using (var redPathXMLFile = XmlReader.Create(mergedXMLPath))
            {
                //Loops thorugh until the EOF
                while (redPathXMLFile.Read())
                {

                    //Searching from the top!!
                    if (redPathXMLFile.IsStartElement())
                    {
                        //Searches for the inner text of the searchTag and returns the value
                        if (redPathXMLFile.ReadToDescendant(searchTag))
                        {
                            redPathXMLFile.Read(); //Reads the next opening or closing tag
                            var result = redPathXMLFile.Value;
                            return result;

                        }

                    }

                }
                return null; //Returns null if the data is not found
            }
        }
        public static string getTagName(string spreadsheetId = "1rkEhkGsitr3VhKayIJpvdoUsIYOPfJUNimMD09CkMuE",string sheetName = "Sheet1",string mergedXMLPath = "C:\\Users\\udhungel\\Downloads\\m@hdcperfagtv004@Platform_Redpath_OVR.xml")
        {
            // string spreadsheetId = spreadsheetIDTextbox.Text; //Impliment this for the input entered in the textbox
            for (int i = 0; i < 6; i++)
            {
                appendEmptyColumn(spreadsheetId);
            }
            
            //Dictionary<string, string> OWMtagNameDict = makeDictionary();
            bool repeatLoop = true;
            List<string> listOfTestNames = new List<string>();
            // The A1 notation of a range to search for a logical table of data.
            // Values will be appended after the last row of the table.
            string range = string.Concat(sheetName, "!A:A"); // TODO: Update placeholder value.

            // The ID of the spreadsheet to update.
            // TODO: Update placeholder value.

            SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = GetCredential(),
                ApplicationName = "RaveReportAutomationUtility",
            });



            SpreadsheetsResource.ValuesResource.GetRequest request =
            sheetsService.Spreadsheets.Values.Get(spreadsheetId, range); //puts an api call to get the items in the row with the provided range

            Data.ValueRange response = request.Execute();

            IList<IList<Object>> sheetCol1Values = response.Values; //generic list to store the row name tags*/



            List<Result_List> values = new List<Result_List>();

            Process_Results.getSPMergeFileCSVFormattedResult(mergedXMLPath,out values);
            List<string> listofScriptTags = new List<string>();
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)//Gets all the tags of the COLUMN A
                {
                  

                    listofScriptTags.Add(row.scriptname);
                }
                int incrementIndex = 0;
                int scriptId = 1;
                foreach (var row in values)//Gets all the tags of the COLUMN A
                {
                    
                    string transactionName = row.transactionname;
                    string scriptName = row.scriptname;
                    string Min = row.min;
                    string Max = row.max;
                    string Avg = row.avg;
                    string Std = row.std;
                    string Count = row.count;
                    string bound1 = row.bound1;
                    string bound2 = row.bound2;
                    string nameTag = row.transactionname.ToString();
                    bool scriptTag = false;
                  

                    if(transactionName == "#Overall Response Time#")
                    {
                        incrementIndex++;
                        scriptTag = true;
                    }
                    int rowIndex = values.IndexOf(row)+incrementIndex;
                    //   var name = row.Select(p => p.ToString()).ToList(); // Converting generic type to string
                    //   string nameTag = string.Join("", name);  // // Converting generic type to string
                  

                    bool loop = true;
                    var startColNum = 1;
                   

                    char rangeConcat = 'A'; //For specifying range
                    do
                    {

                        // Gets the index of the current row where the Metrics is inserted!

                        string nullColumnIndexRange = string.Concat(sheetName,"!"); //Range to be concatinated with

                        string rangeStr = rangeConcat.ToString();
                      
                        string finalRange = string.Concat(nullColumnIndexRange, rangeStr, rowIndex );//Concatinates the string to produce the range string for the right column to enter the date
                     

                        string colData = getColumnData(finalRange);//Loops through the google docs sheet to find the 

                        //Adds the script# title after end of each scripts
                        if (scriptTag == true && rowIndex>1)
                        {
                           // appendSheetFormat(rowIndex - 2);
                            insertNewColumnValue("Script#"+scriptId, null, null, null, null, null, null, null, null, string.Concat(nullColumnIndexRange, rangeStr, rowIndex-1));
                            mergeCells(startColNum-1,rowIndex-2);
                            
                            scriptTag = false;
                            scriptId++;
                        }

                        //Inserts the metrics into the google spreadsheet
                        if (colData == null)
                        {
                            
                            if (rowIndex==1)
                            {
                                //appendSheetFormat(rowIndex);
                                insertNewColumnValue("Script#" + scriptId, null, null, null, null, null, null, null, null, string.Concat(nullColumnIndexRange, rangeStr, rowIndex + 1));
                                mergeCells(startColNum-1, rowIndex);
                                //appendSheetFormat(rowIndex -1);
                                insertSetupInfoColumn(finalRange);
                                scriptId++;
                                incrementIndex++;
                                loop = false;
                            }
                            else
                            {   
                                insertNewColumnValue(transactionName, scriptName, Min, Max, Avg, Std, Count, bound1, bound2, finalRange);
                                loop = false;
                            }


                        }
                        rangeConcat++;
                        startColNum++;

                    } while (loop == true);

                    
                }
            }




            return null;
        }
        public static void rowHeadingTags(int rowIndex, string spreadsheetId = "1rkEhkGsitr3VhKayIJpvdoUsIYOPfJUNimMD09CkMuE", string sheetName = "Sheet1")
        {
            // The A1 notation of a range to search for a logical table of data.
            // Values will be appended after the last row of the table.
            //string range = "1 - Performance Test Results";  // TODO: Update placeholder value.

            // string spreadsheetId = spreadsheetIDTextbox.Text; //Impliment this for the input entered in the textbox

            SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = GetCredential(),
                ApplicationName = "RaveReportAutomationUtility",
            });

            Data.ValueRange valueRange = new Data.ValueRange();
            valueRange.MajorDimension = "ROWS";

            // var output = Math.Round(Convert.ToDouble(getOverallWorkflowMetricsData(nameTag, "Response time[s]", "Avg")), 3);
            var oblist = new List<Object>() { "Min\n(Sec)" };
            oblist.Add("Avg\n(Sec)");
            oblist.Add("Max\n(Sec)");
            oblist.Add("Std");
            oblist.Add("Count");
            oblist.Add("%Request\n(<2 sec)");

            valueRange.Values = new List<IList<object>> { oblist };
            bool loop = true;
            char rangeConcat = 'B';
            do
            {

                // Gets the index of the current row where the Metrics is inserted!

                string nullColumnIndexRange = string.Concat(sheetName,"!"); //Range to be concatinated with

                string rangeStr = rangeConcat.ToString();

                string finalRange = string.Concat(nullColumnIndexRange, rangeStr, rowIndex + 1);//Concatinates the string to produce the range string for the right column to enter the data


                string colData = getColumnData(finalRange);//Loops through the google docs sheet to find the emprty box to store the metrics

                if (colData == null)
                {
                    loop = false;
                    //If an empty column is found the Number is entered in that field.
                    SpreadsheetsResource.ValuesResource.UpdateRequest update = sheetsService.Spreadsheets.Values.Update(valueRange, spreadsheetId, finalRange);
                    update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                    Data.UpdateValuesResponse result = update.Execute();
                }
                rangeConcat++;
            } while (loop == true);




            /*SpreadsheetsResource.ValuesResource.AppendRequest append = sheetsService.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
            append.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            Data.AppendValuesResponse result2 = append.Execute();*/
        }
        public static string getColumnData(string range, string spreadsheetId = "1rkEhkGsitr3VhKayIJpvdoUsIYOPfJUNimMD09CkMuE")
        {
            // The A1 notation of a range to search for a logical table of data.
            // Values will be appended after the last row of the table.
            

            SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = GetCredential(),
                ApplicationName = "RaveReportAutomationUtility",
            });

            SpreadsheetsResource.ValuesResource.GetRequest request =
                   sheetsService.Spreadsheets.Values.Get(spreadsheetId, range);

            Data.ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                // System.Threading.Thread.Sleep(800);
                foreach (var row in values)//Gets all the tags of the ROW A
                {

                    var name = row.Select(p => p.ToString()).ToList(); // Converting generic type to string
                    string colData = string.Join("", name);  // // Converting generic type to string


                    // Print columns A and E, which correspond to indices 0 and 4.
                    return colData;
                }
            }
            return null;

        }

        public static void insertNewScriptEndRow(string scriptName, string Range, string spreadsheetId = "1rkEhkGsitr3VhKayIJpvdoUsIYOPfJUNimMD09CkMuE")
        {

            // The A1 notation of a range to search for a logical table of data.
            // Values will be appended after the last row of the table.
            //string range = "1 - Performance Test Results";  // TODO: Update placeholder value.


            // string spreadsheetId = spreadsheetIDTextbox.Text; //Impliment this for the input entered in the textbox

            SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = GetCredential(),
                ApplicationName = "RaveReportAutomationUtility",
            });

            Data.ValueRange valueRange = new Data.ValueRange();
            valueRange.MajorDimension = "ROWS";


            // var output = Math.Round(Convert.ToDouble(getOverallWorkflowMetricsData(nameTag, "Response time[s]", "Avg")), 3);
            var oblist = new List<Object>() { scriptName };//Try to make
            valueRange.Values = new List<IList<object>> { oblist };


            SpreadsheetsResource.ValuesResource.UpdateRequest update = sheetsService.Spreadsheets.Values.Update(valueRange, spreadsheetId, Range);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            Data.UpdateValuesResponse result = update.Execute();

            /*SpreadsheetsResource.ValuesResource.AppendRequest append = sheetsService.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
            append.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            Data.AppendValuesResponse result2 = append.Execute();*/

        }

        public static void insertNewColumnValue(string transactionName, string scriptName, string Min, string Avg, string Max, string Std, string Count, string Bound1,string Bound2, string Range, string spreadsheetId = "1rkEhkGsitr3VhKayIJpvdoUsIYOPfJUNimMD09CkMuE")
        {

            // The A1 notation of a range to search for a logical table of data.
            // Values will be appended after the last row of the table.
            //string range = "1 - Performance Test Results";  // TODO: Update placeholder value.


            // string spreadsheetId = spreadsheetIDTextbox.Text; //Impliment this for the input entered in the textbox

            SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = GetCredential(),
                ApplicationName = "RaveReportAutomationUtility",
            });

            Data.ValueRange valueRange = new Data.ValueRange();
            valueRange.MajorDimension = "ROWS";


            // var output = Math.Round(Convert.ToDouble(getOverallWorkflowMetricsData(nameTag, "Response time[s]", "Avg")), 3);
            var oblist = new List<Object>() {transactionName,scriptName, Min, Avg, Max, Std, Count, Bound1,Bound2 };//Try to make
            valueRange.Values = new List<IList<object>> { oblist };


            SpreadsheetsResource.ValuesResource.UpdateRequest update = sheetsService.Spreadsheets.Values.Update(valueRange, spreadsheetId, Range);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            Data.UpdateValuesResponse result = update.Execute();

            /*SpreadsheetsResource.ValuesResource.AppendRequest append = sheetsService.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
            append.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            Data.AppendValuesResponse result2 = append.Execute();*/

        }

        //Insert the initial Setup Information:
        public static void insertSetupInfoColumn(string Range, string transactionName= "Transaction Name", string scriptName="Script Name", string Min = "Min", 
            string Avg="Avg", string Max="Max", string Std="Std", string Count="Count", string Bound1="Bound1", string Bound2="Bound2", string spreadsheetId = "1rkEhkGsitr3VhKayIJpvdoUsIYOPfJUNimMD09CkMuE")
        {

            // The A1 notation of a range to search for a logical table of data.
            // Values will be appended after the last row of the table.
            //string range = "1 - Performance Test Results";  // TODO: Update placeholder value.

            // string spreadsheetId = spreadsheetIDTextbox.Text; //Impliment this for the input entered in the textbox

            SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = GetCredential(),
                ApplicationName = "RaveReportAutomationUtility",
            });

            Data.ValueRange valueRange = new Data.ValueRange();
            valueRange.MajorDimension = "ROWS";

            // var output = Math.Round(Convert.ToDouble(getOverallWorkflowMetricsData(nameTag, "Response time[s]", "Avg")), 3);
            var oblist = new List<Object>() { transactionName, scriptName, Min, Avg, Max, Std, Count, Bound1, Bound2 };//Try to make
            valueRange.Values = new List<IList<object>> { oblist };


            SpreadsheetsResource.ValuesResource.UpdateRequest update = sheetsService.Spreadsheets.Values.Update(valueRange, spreadsheetId, Range);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            Data.UpdateValuesResponse result = update.Execute();
        }

        public static Dictionary<string, string> makeDictionary(List<OneClickCore.Result_List> _results)
        {
            Dictionary<string, string> OWMtagName = new Dictionary<string, string>();
            foreach (OneClickCore.Result_List r in _results)
            {
                OWMtagName.Add(r.scriptname, string.Concat(r.scriptname, ".bdf/VUser-Profile1"));
            }
            //OWMtagName.Add("OWM Script #1", "RedPath_Scr01_Login_OnlyMCC_Enabled_03302017.bdf/VUser-Profile1");
            //OWMtagName.Add("OWM Script #2", "RedPath_Scr02_Login_MCC_iMed_Enabled_03302017.bdf/VUser-Profile11");
            //OWMtagName.Add("OWM Script #3", "RedPath_Scr03_Create_Study_Map_eLearning_04072017.bdf/VUser-Profile1");
            //OWMtagName.Add("OWM Script #4", "RedPath_Scr04_Add_Site_To_Study_04112017_2.bdf/VUser-Profile1");
            //OWMtagName.Add("OWM Script #5", "RedPath_Scr05_Add_New_User_To_Study_04072017.bdf/VUser-Profile1");
            //OWMtagName.Add("OWM Script #6", "RedPath_Scr06_Complete_eLearning_Course_04202017.bdf/VUser-Profile1");
            //OWMtagName.Add("OWM Script #7", "RedPath_Scr07_Edit_Study_CDSite_StdSite.bdf/VUser-Profile1");
            //OWMtagName.Add("OWM Script #9", "RedPath_Scr09_Navigate_Users_Sites_04192017.bdf/VUser-Profile1");
            //OWMtagName.Add("OWM Script #10", "RedPath_Scr10_2_Delete_Add_Roles_To_User_04122017.bdf/VUser-Profile1");
            //OWMtagName.Add("OWM Script #11", "RedPath_Scr11_Add_Additional_Role_To_Study_04172017.bdf/VUser-Profile1");
            //OWMtagName.Add("OWM Script #12", "RedPath_Scr12_Edit_Roles_04242014.bdf/VUser-Profile1");
            //OWMtagName.Add("OWM Script #13", "RedPath_Scr13_AddClientDivision_User.bdf/VUser-Profile1");
            //OWMtagName.Add("OWM Script #14", "RedPath_Scr14_CreateEnvironment_AssignUser.bdf/VUser-Profile1");

            return OWMtagName;
        }
        public static string getOWMTagName(string nameTag, Dictionary<string, string> OWMTagNameDict)
        {
            return OWMTagNameDict[nameTag];
        }
        public static void appendEmptyColumn(string spreadsheetId = "1rkEhkGsitr3VhKayIJpvdoUsIYOPfJUNimMD09CkMuE", int sheetId = 0)
        {

            // The A1 notation of a range to search for a logical table of data.
            // Values will be appended after the last row of the table.
            //   string range = "1 - Performance Test Results";  // TODO: Update placeholder value.
            //"'1 - Performance Test Results'!A:G"
            // The ID of the spreadsheet to update.

            SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = GetCredential(),
                ApplicationName = "RaveReportAutomationUtility",
            });

            Data.Request reqBody = new Data.Request
            {
                AppendDimension = new Data.AppendDimensionRequest
                {
                    SheetId = sheetId,
                    Dimension = "COLUMNS",
                    Length = 1

                }
            };

            List<Data.Request> requests = new List<Data.Request>();

            requests.Add(reqBody);


            // TODO: Assign values to desired properties of `requestBody`:
            Data.BatchUpdateSpreadsheetRequest requestBody = new Data.BatchUpdateSpreadsheetRequest();
            requestBody.Requests = requests;

            SpreadsheetsResource.BatchUpdateRequest request = sheetsService.Spreadsheets.BatchUpdate(requestBody, spreadsheetId);
            //SpreadsheetsResource.BatchUpdateRequest request = sheetsService.Spreadsheets.BatchUpdate(requestBody, spreadsheetId);


            // To execute asynchronously in an async method, replace `request.Execute()` as shown:
            Data.BatchUpdateSpreadsheetResponse response = request.Execute();
            // Data.BatchUpdateSpreadsheetResponse response = await request.ExecuteAsync();

            // TODO: Change code below to process the `response` object:
            //Console.WriteLine(JsonConvert.SerializeObject(response));


            //     SpreadsheetsResource.ValuesResource.AppendRequest request =
            //         sheetsService.Spreadsheets.Values.Append(body, spreadsheetId, range);
            //   request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
            //request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
            //var response = request.Execute();            
        }
        public static void mergeCells(int startCol,int startRow, string spreadsheetId = "1rkEhkGsitr3VhKayIJpvdoUsIYOPfJUNimMD09CkMuE", int sheetId = 0)
        {
            // The A1 notation of a range to search for a logical table of data.
            // Values will be appended after the last row of the table.

            //"'1 - Performance Test Results'!A:G"
            // The ID of the spreadsheet to update.


            // string spreadsheetId = spreadsheetIDTextbox.Text; //Impliment this for the input entered in the textbox



            SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = GetCredential(),
                ApplicationName = "RaveReportAutomationUtility",
            });

            Data.Request reqBody = new Data.Request
            {
                MergeCells = new Data.MergeCellsRequest
                {
                    Range = new Data.GridRange
                    {
                        SheetId = sheetId,
                        StartRowIndex = startRow,
                        EndRowIndex = startRow+1,
                        StartColumnIndex = startCol,
                        EndColumnIndex = startCol + 9


                    },
                    MergeType = "MERGE_ROWS"

                }
            };

            List<Data.Request> requests = new List<Data.Request>();

            requests.Add(reqBody);


            // TODO: Assign values to desired properties of `requestBody`:
            Data.BatchUpdateSpreadsheetRequest requestBody = new Data.BatchUpdateSpreadsheetRequest();
            requestBody.Requests = requests;

            SpreadsheetsResource.BatchUpdateRequest request = sheetsService.Spreadsheets.BatchUpdate(requestBody, spreadsheetId);
            //SpreadsheetsResource.BatchUpdateRequest request = sheetsService.Spreadsheets.BatchUpdate(requestBody, spreadsheetId);


            // To execute asynchronously in an async method, replace `request.Execute()` as shown:
            Data.BatchUpdateSpreadsheetResponse response = request.Execute();
            // Data.BatchUpdateSpreadsheetResponse response = await request.ExecuteAsync();

            // TODO: Change code below to process the `response` object:
            //Console.WriteLine(JsonConvert.SerializeObject(response));


            //     SpreadsheetsResource.ValuesResource.AppendRequest request =
            //         sheetsService.Spreadsheets.Values.Append(body, spreadsheetId, range);
            //   request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
            //request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
            //var response = request.Execute();

        }


        /*    public static void appendSheetFormat( int startRow,string spreadsheetId = "1rkEhkGsitr3VhKayIJpvdoUsIYOPfJUNimMD09CkMuE", int sheetId = 0)
       {

           // The A1 notation of a range to search for a logical table of data.
           // Values will be appended after the last row of the table.
           //   string range = "1 - Performance Test Results";  // TODO: Update placeholder value.
           //"'1 - Performance Test Results'!A:G"
           // The ID of the spreadsheet to update.

           SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
           {
               HttpClientInitializer = GetCredential(),
               ApplicationName = "RaveReportAutomationUtility",
           });


           //request to make the text bold in the provided range in the spreadsheet
           Data.Request reqBody = new Data.Request
           {
               RepeatCell = new Data.RepeatCellRequest
               {
                   Fields = "*",
                   Range = new Data.GridRange
                   {
                       SheetId = sheetId,
                       StartRowIndex = startRow,
                       EndRowIndex = startRow +1
                   },
                   Cell = new Data.CellData
                   {
                      UserEnteredFormat = new Data.CellFormat
                      {

                          TextFormat = new Data.TextFormat
                          {

                              Bold= true
                          }
                      }
                   }

               },


           };
           List<Data.Request> requests = new List<Data.Request>();
           requests.Add(reqBody);

           //Increases the width of the columns 1 and 2 to fit the text in the spreadsheet
          /* if(startRow == 1)
           {
               Data.Request reqBody1 = new Data.Request
               {

                   UpdateDimensionProperties = new Data.UpdateDimensionPropertiesRequest
                   {
                       Fields = "*",
                       Range = new Data.DimensionRange
                       {
                           SheetId = sheetId,
                           StartIndex = 0,
                           EndIndex = 2,
                           Dimension = "COLUMNS"
                       },
                       Properties = new Data.DimensionProperties
                       {
                           PixelSize = 300
                       }

                   },
               };
               requests.Add(reqBody1);
           }






           // TODO: Assign values to desired properties of `requestBody`:

           Data.BatchUpdateSpreadsheetRequest requestBody = new Data.BatchUpdateSpreadsheetRequest();
           requestBody.Requests = requests;

           SpreadsheetsResource.BatchUpdateRequest request = sheetsService.Spreadsheets.BatchUpdate(requestBody, spreadsheetId);
           //SpreadsheetsResource.BatchUpdateRequest request = sheetsService.Spreadsheets.BatchUpdate(requestBody, spreadsheetId);


           // To execute asynchronously in an async method, replace `request.Execute()` as shown:
           Data.BatchUpdateSpreadsheetResponse response = request.Execute();
           // Data.BatchUpdateSpreadsheetResponse response = await request.ExecuteAsync();

           // TODO: Change code below to process the `response` object:
           //Console.WriteLine(JsonConvert.SerializeObject(response));


           //     SpreadsheetsResource.ValuesResource.AppendRequest request =
           //         sheetsService.Spreadsheets.Values.Append(body, spreadsheetId, range);
           //   request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
           //request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
           //var response = request.Execute();            
       }*/



        public static UserCredential GetCredential()
        {
            UserCredential credential;
            string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
            credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");
            string[] Scopes = { SheetsService.Scope.Drive };

            var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read);

            credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,

                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
            //stream.Close();
            return credential;
        }
    }


    public class apiCall
    {

    }
}

