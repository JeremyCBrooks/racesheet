using Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Collections.Generic;
using System;
using System.Threading;

namespace racesheet
{
    class Program
    {
        class Heat
        {
            public List<string> Racers { get; set; } = new List<string>();
        }

        static void Main(string[] args)
        {
            string workbookRequested = "RaceSheet.xlsx";

            Application oXL = null;
            _Workbook oWB = null;

            try
            {
                Console.Out.WriteLine("Searching for running Excel instance...");
                oXL = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application") as Application;
                if (null == oXL)
                {
                    throw new Exception($"Could not find Excel on this system. Make sure it is installed.");
                }
                Console.Out.WriteLine($"done");

                Console.Out.WriteLine($"Searching for {workbookRequested}...");
                foreach (Workbook workbook in oXL.Application.Workbooks)
                {
                    if (workbook.Name == workbookRequested)
                    {
                        oWB = workbook;
                        break;
                    }
                }
                Console.Out.WriteLine($"done");

                if (null == oWB)
                {
                    throw new Exception($"Could not find a running instance of {workbookRequested}. Make sure the spreadsheet is open and editable.");
                }

                Console.Out.WriteLine($"Looking for config values...");
                _Worksheet configWS = oWB.Worksheets["Config"];
                int heatsPerRacer = int.Parse(configWS.Cells[2, 3].Value.ToString());
                int lanes = int.Parse(configWS.Cells[2, 5].Value.ToString());
                string dropWorstTimeVal = (string)(configWS.Cells[2, 7].Value.ToString()).ToLowerInvariant().Trim();
                bool dropWorstTime = (dropWorstTimeVal == "1" || dropWorstTimeVal == "true" || dropWorstTimeVal == "yes");

                Console.Out.WriteLine($"Heats Per Racer: {heatsPerRacer}");
                Console.Out.WriteLine($"Lanes: {lanes}");
                Console.Out.WriteLine($"Drop Worst Time: {dropWorstTime}");

                int totalHeatRows = 0;

                {
                    var racers = new List<string>();

                    Console.Out.WriteLine($"Collecting roster...");
                    _Worksheet rosterWS = oWB.Worksheets["Roster"];
                    foreach (Range row in rosterWS.UsedRange.Rows)
                    {
                        string name = row.Cells[2, 1].Value;

                        if (null != name && !racers.Any(r => r == name))
                        {
                            racers.Add(name);
                        }
                    }
                    Console.Out.WriteLine($"done");

                    _Worksheet heatsWS = oWB.Worksheets["Heats"];

                    //clear existing heats
                    Console.Out.WriteLine($"Clearing existing heats...");
                    heatsWS.Range["A2", "E500"].Value2 = "";
                    Console.Out.WriteLine($"done");

                    //calculate new heats
                    Console.Out.WriteLine($"Calculating new heats...");
                    int count = 0;
                    for (int roundI = 0; roundI < heatsPerRacer; ++roundI)
                    {
                        racers.Shuffle();
                        for (int racerI = 0; racerI < racers.Count; ++racerI, ++count)
                        {
                            int lane = (count % lanes);
                            int heatNumber = count / lanes;
                            var racer = racers[racerI];

                            int row = count + 2;
                            heatsWS.Cells[row, 1] = heatNumber + 1;
                            heatsWS.Cells[row, 2] = lane + 1;
                            heatsWS.Cells[row, 3] = racer;
                            heatsWS.Cells[row, 4] = $"=VLOOKUP(C{row},Roster!$A:$B,2,FALSE)";
                            heatsWS.Cells[row, 5] = "";
                        }
                    }

                    //generate BYEs
                    Console.Out.WriteLine($"Generating BYEs...");
                    for (; count % lanes != 0; ++count)
                    {
                        int heatNumber = count / lanes;
                        int lane = (count % lanes) + 1;
                        int row = count + 2;
                        heatsWS.Cells[row, 1] = heatNumber + 1;
                        heatsWS.Cells[row, 2] = lane;
                        heatsWS.Cells[row, 3] = "BYE";
                        heatsWS.Cells[row, 4] = "";
                        heatsWS.Cells[row, 5] = "";
                    }
                    Console.Out.WriteLine($"done");

                    totalHeatRows = count;
                }

                //reset standings
                {
                    var racers = new List<string>();

                    Console.Out.WriteLine($"Collecting roster...");
                    _Worksheet rosterWS = oWB.Worksheets["Roster"];
                    foreach (Range row in rosterWS.UsedRange.Rows)
                    {
                        string name = row.Cells[2, 1].Value;

                        if (null != name && !racers.Any(r => r == name))
                        {
                            racers.Add(name);
                        }
                    }
                    _Worksheet standingsWS = oWB.Worksheets["Standings"];
                    standingsWS.Range["A2", "C500"].Value2 = "";
                    Console.Out.WriteLine("Reset standings...");
                    for (int racerI = 0; racerI < racers.Count; ++racerI)
                    {
                        var racer = racers[racerI];

                        int row = racerI + 2;
                        standingsWS.Cells[row, 1] = racer;
                        standingsWS.Cells[row, 2] = $"=VLOOKUP(A{row},Roster!$A:$B,2,FALSE)";

                        var timeCalc = $"=SUMIFS(Heats!$E$2:$E$500,Heats!$C$2:$C$500,A{row},Heats!$D$2:$D$500,B{row})";
                        if (dropWorstTime)
                        {
                            timeCalc += $"-IF(COUNTIFS(Heats!$E$2:$E$500, \"<>\", Heats!$C$2:$C$500,A{row},Heats!$D$2:$D$500,B{row})=Config!$C$2, MAXIFS(Heats!$E$2:$E$500,Heats!$C$2:$C$500,A{row},Heats!$D$2:$D$500,B{row}), 0)";
                        }
                        standingsWS.Cells[row, 3] = timeCalc;
                    }
                    Console.Out.WriteLine($"done");
                }
            }
            catch (Exception ex)
            {
                Console.Out.WriteLine(ex.Message);
                Console.ReadKey();
            }
            finally
            {
                if(null != oXL)
                {
                    try
                    {
                        while (System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL) > 0) ;
                    }
                    catch { }
                    finally
                    {
                        oXL = null;
                    }
                }
            }
        }
    }

    public static class ThreadSafeRandom
    {
        [ThreadStatic] private static Random Local;

        public static Random Random
        {
            get { return Local ?? (Local = new Random(unchecked(Environment.TickCount * 31 + Thread.CurrentThread.ManagedThreadId))); }
        }
    }

    static class MyExtensions
    {
        public static void Shuffle<T>(this IList<T> list)
        {
            int n = list.Count;
            while (n > 1)
            {
                n--;
                int k = ThreadSafeRandom.Random.Next(n + 1);
                T value = list[k];
                list[k] = list[n];
                list[n] = value;
            }
        }
    }
}
