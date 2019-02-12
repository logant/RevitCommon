using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
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
                // It would have shown a random, though curated image and a haiku from Basho. This would happen for about 15% of use, 
                // and then 5% would get an additional pop-up notification saying money has been deducted from the project and given 
                // to LINE. The rest of my colleagues did not find it as amusing as I did so it is disabled, though preserved. :(
                EasterEgg();
            }
            catch { } // FileUtils failed, but as it doesn't affect the user we'll ignore
        }

        /// <summary>
        /// This was just for my amusement. I made a single tool that would occasionally pop up a message about money being deducted
        /// and given to LINE, which some people found amusing. It was a little used tool, so when applied to a larger group there was
        /// some push back against having jokes, so I've disabled them for now by setting the easter egg percentage to 0 (Variable at
        /// RevitCommon/paths/ee-pct in the RevitCommon.config file) and as the default if no config parameter is found.
        /// </summary>
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
                bool lineNotification = choice <= 50;

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

        /// <summary>
        /// This is related to the EasterEgg method. It shows a message from the Revit Info Center (top right on title bar).
        /// </summary>
        /// <param name="title">Title of the message</param>
        /// <param name="message">Main message content</param>
        /// <param name="toolTip">Tooltip text if desired.</param>
        public static void ShowBalloonTip(string title, string message, string toolTip)
        {
            Autodesk.Internal.InfoCenter.ResultItem ri = new Autodesk.Internal.InfoCenter.ResultItem
            {
                Category = title,
                Title = message,
                TooltipText = toolTip,
                Uri = new System.Uri("http://www.hksline.com"),
                IsFavorite = true,
                IsNew = true
            };

            ri.ResultClicked += new EventHandler<Autodesk.Internal.InfoCenter.ResultClickEventArgs>(ri_ResultClicked);

            adWin.ComponentManager.InfoCenterPaletteManager.ShowBalloon(ri);
        }

        /// <summary>
        /// This is also related to the EasterEgg method. This handles the InfoCenter click event.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void ri_ResultClicked(object sender, Autodesk.Internal.InfoCenter.ResultClickEventArgs e)
        {
            Autodesk.Internal.InfoCenter.ResultItem ri = (Autodesk.Internal.InfoCenter.ResultItem)sender;
            System.Diagnostics.Process.Start(ri.Uri.ToString());
        }

        /// <summary>
        /// Helper method for reading the RevitCommon.config file. This takes an XML Node path and returns
        /// the first path that satisfies it. It's intended for the RevitCommon/paths section of the file
        /// and will attempt to build a valid file path.
        /// </summary>
        /// <param name="xmlNodePath"></param>
        /// <returns></returns>
        public static string GetPath(string xmlNodePath)
        {
            string path = string.Empty;
            var directoryInfo = new FileInfo(typeof(FileUtils).Assembly.Location).Directory;
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

        /// <summary>
        /// This method will attempt to get plugin data from the RevitCommon.config file. This is for the plugins I've made that are
        /// released publicly via github (basically what I created at LMN Architects). This way I can have them default to the HKS tab
        /// for deploying locally, and anyone else that downloads them can set them to their own Tab/Panel locations on the Ribbon.
        /// </summary>
        /// <param name="pluginName"></param>
        /// <param name="helpPath"></param>
        /// <param name="tabName"></param>
        /// <param name="panelName"></param>
        /// <returns></returns>
        public static bool GetPluginSettings(string pluginName, out string helpPath, out string tabName, out string panelName)
        {
            helpPath = null;
            tabName = null;
            panelName = null;
            string configPath = new FileInfo(typeof(FileUtils).Assembly.Location).Directory.FullName + "\\RevitCommon.config";
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
            {
                if (File.Exists(helpNode.InnerText))
                    helpPath = helpNode.InnerText;
                else
                {
                    var combPath = Path.Combine(Path.GetDirectoryName(typeof(FileUtils).Assembly.Location), helpNode.InnerText);
                    var uriPath = Path.GetFullPath((new Uri(combPath)).LocalPath);
                    if (!string.IsNullOrWhiteSpace(uriPath))
                        helpPath = uriPath;
                }
            }

            if (tabNode != null)
                tabName = tabNode.InnerText;
            if (panelNode != null)
                panelName = panelNode.InnerText;

            return true;
        }

        /*

        ===============================================================================================================================

        I BELIEVE THIS METHOD IS UNNECESSARY AS IT'S JUST A LONGER WAY DOING WHAT SYSTEM.IO.PATH.COMBINE DOES WILL VERIFY IT STILL WORKS
        AND THEN SEE ABOUT DELETING IT BEFORE PUSHING IT BACK TO THE REPO.

        ===============================================================================================================================

        /// <summary>
        /// This is used to find the full path when given a relative path. It's intended to be used for things like
        /// the RevitCommon/plugin/help-path that's part of a plugin definition in the RevitCommon.config file. It combines
        /// the relative path in the config file with the assembly's location.
        /// </summary>
        /// <param name="initPath"></param>
        /// <returns></returns>
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
        */
    }
}
