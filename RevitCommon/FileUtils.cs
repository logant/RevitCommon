using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using Autodesk.Revit.DB;
using adWin = Autodesk.Windows;

namespace RevitCommon
{
    public class FileUtils
    {
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
                // Check to see if the log file exists
                string logPath = GetPath("RevitCommon/paths/log-path");
                if (string.IsNullOrEmpty(logPath))
                    return;

                // Get the current year so app usage is organized into different files by year
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
                // else it probably is being run outside of network or the file is busy.

                // I thought this would be an amusing joke, popping up a random balloon message from the Revit communication center
                // It would have shown a random, though currated, image and a haiku from Basho. This would happen for about 15% of use, 
                // and then 5% would get an additional pop-up notification saying money has been deducted from the project and given 
                // to LINE. The rest of my colleagues did not find it as amusing as I did so it is disabled, though preserved. :(
                EasterEgg();
            }
            catch { } // FileUtils failed, but as it doesn't affect the user we'll ignore
        }

        private static void EasterEgg()
        {
            // EasterEgg path
            string path = GetPath("RevitCommon/paths/ee-path");
            if (string.IsNullOrWhiteSpace(path))
                return;
            if (!File.Exists(path))
                return;

            string pctStr = GetPath("RevitCommon/paths/ee-pct");
            if (!double.TryParse(pctStr, out double pct))
                pct = 0;

            // EasterEgg Pop-Up
            Random r = new Random();
            int choice = r.Next(1000);

            // Currently disabled
            int limit = Convert.ToInt32(1000.0 * pct);

            if (choice < limit)
            {
                bool lineNotification = false;
                if (choice <= 50)
                    lineNotification = true;
                string message = "$10 has been deducted from your project budget and given to HKS LINE to fund more awesome tools!";

                if (File.Exists(path))
                {
                    List<string> bashoHaikus = new List<string>();
                    StreamReader reader = new StreamReader(path);
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

        public static string GetPath(string xmlNodePath)
        {
            string path = string.Empty;
            var directoryInfo = new FileInfo(typeof(HKS).Assembly.Location).Directory;
            if (directoryInfo == null)
                return null;

            var configPath = directoryInfo.FullName + "\\RevitCommon.config";
            if (!File.Exists(configPath))
                return null;

            var configStr = File.ReadAllText(configPath);
            XmlDocument xDoc = new XmlDocument();
            try
            {
                xDoc.LoadXml(configStr);
                XmlNode node = xDoc.SelectSingleNode(xmlNodePath);
                path = node.InnerText;

                // Check to see if the path is a directory
                if (!Directory.Exists(path))
                {
                    // Check to see if it's a number
                    if (!double.TryParse(path, out double numCheck))
                    {
                        // Check to see if it is instead a file
                        var fileDir = new FileInfo(path).Directory;
                        if ((fileDir != null && !Directory.Exists(fileDir.FullName)) || fileDir == null)
                            return null;
                    }
                }
            }
            catch
            {
                return null;
            }

            return path;
        }

        public static bool GetPluginSettings(string pluginName, out string helpPath, out string tabName, out string panelName)
        {
            helpPath = null;
            tabName = null;
            panelName = null;
            string configPath = new FileInfo(typeof(HKS).Assembly.Location).Directory.FullName + "\\RevitCommon.config";
            if (!File.Exists(configPath))
            {
                return false;
            }

            // Load the XmlDocument that contains the settings
            XmlDocument xDoc = new XmlDocument();
            try
            {
                xDoc.Load(configPath);
            }
            catch
            {
                return false;
            }

            // Look for plugin nodes
            XmlNodeList nodes = xDoc.SelectNodes("RevitCommon/plugin");
            XmlNode pNode = null;
            foreach (XmlNode n in nodes)
            {
                if (n.Attributes != null && n.Attributes["name"].Value == pluginName)
                {
                    pNode = n;
                    break;
                }
            }

            if (pNode == null)
                return false;

            XmlNode helpNode = pNode.SelectSingleNode("help-path");
            XmlNode tabNode = pNode.SelectSingleNode("tab-name");
            XmlNode panelNode = pNode.SelectSingleNode("panel-name");

            if (helpNode != null)
                helpPath = GetFullPath(helpNode.InnerText);
            if (tabNode != null)
                tabName = tabNode.InnerText;
            if (panelNode != null)
                panelName = panelNode.InnerText;

            return true;
        }

        private static string GetFullPath(string initPath)
        {
            // Check to see if the directory is an absolute path
            if (File.Exists(initPath))
                return initPath;

            // Check to see if it's an ancestor path to where we are at
            string loc = typeof(FileUtils).Assembly.Location;
            DirectoryInfo dirInfo = new FileInfo(loc).Directory;

            // Check to see if the file is in the same directory
            //if (File.Exists(dirInfo.FullName + "\\" + initPath))
//                return dirInfo.FullName + "\\" + initPath;
            
            // Check to see if there are some steps back we need to take.
            string[] pathParts = initPath.Split(new char[] { '\\' });
            if (pathParts.Length == 0)
                return null;

            string combined = dirInfo.FullName + "\\" + initPath;
            if (pathParts[0] != "..")
            {
                combined = dirInfo.FullName + "\\" + initPath;
                if (File.Exists(combined))
                    return combined;

                // File doesn't seem to exist
                return null;
            }





            // find out how many steps back we need to make
            int stepsBack = 0;
            foreach (string part in pathParts)
            {
                if (part == "..")
                    stepsBack++;
            }

            // Check get the new base location
            for (int i = 0; i < stepsBack; i++)
            {
                dirInfo = dirInfo.Parent;
            }
            // get the initpath without the steps back
            string newInit = string.Empty;
            for (int i = stepsBack; i < pathParts.Length; i++)
            {
                newInit += "\\" + pathParts[i];
            }

            // Final Path
            combined = dirInfo.FullName + newInit;
            if(File.Exists(combined))
                return combined;

            return null;
        }
    }
}
