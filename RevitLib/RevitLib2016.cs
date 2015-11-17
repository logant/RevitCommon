using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using adWin = Autodesk.Windows;

namespace LINE
{
    public class RevitLib2016
    {
        /// <summary>
        /// Add buttons to the Revit UI
        /// </summary>
        /// <param name="revApp">Revit application for button generation</param>
        /// <param name="tabName">Name of tab to place buttons</param>
        /// <param name="panelName">Name of panel to place buttons</param>
        /// <param name="buttons">List of PushButtonData to generate the buttons</param>
        /// <returns></returns>
        public static bool AddToRibbon(UIControlledApplication revApp, string tabName, string panelName, List<PushButtonData> buttons)
        {
            try
            {
                // Verify if the tab exists, create it if ncessary
                adWin.RibbonControl ribbon = adWin.ComponentManager.Ribbon;
                adWin.RibbonTab tab = null;
                foreach(adWin.RibbonTab t in ribbon.Tabs)
                {
                    if (t.Id == tabName)
                    {
                        tab = t;
                        break;
                    }
                }
                if (tab == null)
                    revApp.CreateRibbonTab(tabName);

                // Verify if the panel exists
                List<RibbonPanel> panels = revApp.GetRibbonPanels(tabName);
                RibbonPanel panel = null;
                foreach (RibbonPanel rp in panels)
                {
                    if (rp.Name == panelName)
                    {
                        panel = rp;
                        break;
                    }
                }
                if (panel == null)
                    panel = revApp.CreateRibbonPanel(tabName, panelName);
                
                // Add the button(s) to the panel
                foreach(PushButtonData pbd in buttons)
                {
                    panel.AddItem(pbd);
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

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
    }
}
