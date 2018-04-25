using System;
using System.Collections.Generic;
using System.IO;
using Autodesk.Revit.DB;
using adWin = Autodesk.Windows;

namespace RevitCommon
{
    public class HKS
    {
        // Old Log Path. The line server was moved from NT11 to LINE-FS-01
        //private static string logPath = @"\\nt11\00\00603.000\01_LINE\tlogan\experiements\LINEtools_[YEAR].txt";

        // Current server path.
        private static string logPath = @"\\line-fs-01\01_Projects\Side-Projects\RevitPlugin_Logging\LINEtools_[YEAR].txt";

        /// <summary>
        /// Returns the scale of a document assuming a meter as the base unit.  So meters should be 1.0, feet should be 0.3048 and so on.
        /// </summary>
        /// <param name="doc">Revit document to get the scale</param>
        /// <returns>Scale of the units compared to meters.</returns>
        public static double RevitScaleToMeters(Document doc)
        {
            double scale = 1.0;

            Units units = doc.GetUnits();
            FormatOptions fo = units.GetFormatOptions(UnitType.UT_Length);
            DisplayUnitType dut = fo.DisplayUnits;
            switch (dut)
            {
                case DisplayUnitType.DUT_CENTIMETERS:
                    scale = 0.01;
                    break;
                case DisplayUnitType.DUT_DECIMAL_FEET:
                    scale = 0.3048;
                    break;
                case DisplayUnitType.DUT_DECIMAL_INCHES:
                    scale = 0.0254;
                    break;
                case DisplayUnitType.DUT_DECIMETERS:
                    scale = 0.1;
                    break;
                case DisplayUnitType.DUT_FEET_FRACTIONAL_INCHES:
                    scale = 0.3048;
                    break;
                case DisplayUnitType.DUT_FRACTIONAL_INCHES:
                    scale = 0.0254;
                    break;
                case DisplayUnitType.DUT_MILLIMETERS:
                    scale = 0.001;
                    break;
                case DisplayUnitType.DUT_METERS:
                    scale = 1.0;
                    break;
                case DisplayUnitType.DUT_METERS_CENTIMETERS:
                    scale = 1.0;
                    break;
                default:
                    scale = 1.0;
                    break;
            }

            return scale;
        }

        /// <summary>
        /// Log usage of in-house commands to a CSV file
        /// </summary>
        /// <param name="commandName">Name of the command used</param>
        /// <param name="appVersion">Revit version</param>
        /// <param name="userName">User who used the command</param>
        public static void WriteToHome(string commandName, string appVersion, string userName)
        {
            try
            {
                // Get the current year so app usage is organized into different fiels by year
                string year = DateTime.Now.Year.ToString();

                // Path to where the log is stored
                string userLogFilePath = logPath.Replace("[YEAR]", year);

                // If the file already exists, just append a new line to the end of the file.
                if (File.Exists(userLogFilePath))
                {
                    using (StreamWriter sw = File.AppendText(userLogFilePath))
                    {
                        sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "," + userName + "," + commandName + "," + appVersion);
                    }
                }
                // If the file doesn't exist but the directory does, create the file and add the one line to it.
                else if (new FileInfo(userLogFilePath).Directory.Exists)
                {
                    string[] newData = { "DATE & TIME, USERNAME, PLUGIN NAME, APPLICATION", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "," + userName + "," + commandName + "," + appVersion };
                    File.WriteAllLines(userLogFilePath, newData);
                }
                // else it probably is being run outside of HKS's network or the file is busy.  Consider making a simple webcall to a website?

                // EasterEgg Pop-Up
                Random r = new Random();
                int choice = r.Next(1000);
                // Currently disabled
                choice = 1001;
                if (choice <= 150)
                {
                    bool lineNotification = false;
                    if (choice <= 50)
                        lineNotification = true;
                    string message = "$10 has been deducted from your project budget and given to HKS LINE to fund more awesome tools!";

                    if (File.Exists(@"\\nt11\00\00603.000\01_LINE\tlogan\_PluginContent\basho.txt"))
                    {
                        List<string> bashoHaikus = new List<string>();
                        StreamReader reader = new StreamReader(@"\\nt11\00\00603.000\01_LINE\tlogan\_PluginContent\basho.txt");
                        string line;
                        while ((line = reader.ReadLine()) != null)
                        {
                            bashoHaikus.Add(line.Replace("\\n", Environment.NewLine));
                        }

                        choice = r.Next(10);
                        if (choice > 6)
                        {
                            choice = r.Next(bashoHaikus.Count - 1);
                            message = bashoHaikus[choice];
                        }
                    }

                    // Show the apprpriate notification.
                    if (lineNotification)
                        UI.Notification("LINE Thanks You!", message, true);
                    else
                        ShowBalloonTip("LINE Thanks You!", message, string.Empty);
                }
            }
            catch 
            {
            }
        }

        public static void ShowBalloonTip(string title, string message, string toolTip)
        {
            Autodesk.Internal.InfoCenter.ResultItem ri = new Autodesk.Internal.InfoCenter.ResultItem();
            ri.Category = title;
            ri.Title = message;
            ri.TooltipText = toolTip;
            ri.Uri = new System.Uri("http://www.hksline.com");

            ri.IsFavorite = true;
            ri.IsNew = true;
            ri.ResultClicked += new EventHandler<Autodesk.Internal.InfoCenter.ResultClickEventArgs>(ri_ResultClicked);

            adWin.ComponentManager.InfoCenterPaletteManager.ShowBalloon(ri);
        }

        private static void ri_ResultClicked(object sender, Autodesk.Internal.InfoCenter.ResultClickEventArgs e)
        {
            Autodesk.Internal.InfoCenter.ResultItem ri = (Autodesk.Internal.InfoCenter.ResultItem)sender;
            System.Diagnostics.Process.Start(ri.Uri.ToString());
        }
    }
}
