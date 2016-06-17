using System.Collections.Generic;
using Autodesk.Revit.UI;
using adWin = Autodesk.Windows;

namespace RevitCommon
{
    public class UI
    {
        /// <summary>
        /// Add one or more buttons to the Revit UI
        /// </summary>
        /// <param name="revApp">Revit application for button generation</param>
        /// <param name="tabName">Name of tab to place buttons</param>
        /// <param name="panelName">Name of panel to place buttons</param>
        /// <param name="buttons">List of PushButtonData to generate the buttons</param>
        /// <returns></returns>
        public static bool AddToRibbon(UIControlledApplication revApp, string tabName, string panelName, List<PushButtonData> buttons)
        {
            RibbonPanel panel = GetRibbonPanel(revApp, tabName, panelName);

            // Add the button(s) to the panel
            if (panel != null)
            {
                foreach (PushButtonData pbd in buttons)
                {
                    panel.AddItem(pbd);
                }

                return true;
            }
            else
            {
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                foreach(PushButtonData pbd in buttons)
                {
                    sb.AppendLine(pbd.Text);
                }
                TaskDialog.Show("Error", "Could not add split button to the Revit ribbon for:\n" + sb.ToString());
                return false;
            }
        }

        /// <summary>
        /// Add a single button to the Revit UI
        /// </summary>
        /// <param name="revApp">Revit application for button generation</param>
        /// <param name="tabName">Name of tab to place buttons</param>
        /// <param name="panelName">Name of panel to place buttons</param>
        /// <param name="button">List of PushButtonData to generate the buttons</param>
        /// <returns></returns>
        public static bool AddToRibbon(UIControlledApplication revApp, string tabName, string panelName, PushButtonData button)
        {
            RibbonPanel panel = GetRibbonPanel(revApp, tabName, panelName);

            // Add the button to the panel
            if (panel != null)
            {
                panel.AddItem(button);
                return true;
            }
            else
            {
                TaskDialog.Show("Error", "Could not add split button to the Revit ribbon for:\n" + button.Text);
                return false;
            }

        }

        /// <summary>
        /// Add a SplitPushButton to Revit
        /// </summary>
        /// <param name="revApp">Revit's UIControlledApplication for adding the button</param>
        /// <param name="tabName">Name of the tab you want to add the button to.</param>
        /// <param name="panelName">Name of the panel on the tab you want to add the button</param>
        /// <param name="button">SplitPushButton to add</param>
        /// <returns></returns>
        public static SplitButton AddToRibbon(UIControlledApplication revApp, string tabName, string panelName, SplitButtonData button)
        {
            RibbonPanel panel = GetRibbonPanel(revApp, tabName, panelName);

            // Add the button to the panel
            if (panel != null)
                return panel.AddItem(button) as SplitButton;
            else
            {
                TaskDialog.Show("Error", "Could not add split button to the Revit ribbon for:\n" + button.Text);
                return null;
            }
                

        }

        /// <summary>
        /// Create a stack of 3 buttons in the Revit UI.
        /// </summary>
        /// <param name="revApp">Revit's UIControlledApplication for adding the button</param>
        /// <param name="tabName">Name of the tab you want to add the button to.</param>
        /// <param name="panelName">Name of the panel on the tab you want to add the button</param>
        /// <param name="button0">Top button in the stack</param>
        /// <param name="button1">Middle button in the stack</param>
        /// <param name="button2">Bottom button in the stack</param>
        public static bool AddToRibbon(UIControlledApplication revApp, string tabName, string panelName, PushButtonData button0, PushButtonData button1, PushButtonData button2)
        {
            RibbonPanel panel = GetRibbonPanel(revApp, tabName, panelName);

            // Add the button to the panel
            if (panel != null)
                panel.AddStackedItems(button0, button1, button2);
            else
                TaskDialog.Show("Error", "Could not create stacked buttons for:\n" + button0.Text + "\n" + button1.Text + "\n" + button2.Text);

            return true;
        }

        private static RibbonPanel GetRibbonPanel(UIControlledApplication revApp, string tabName, string panelName)
        {
            try
            {
                // Verify if the tab exists, create it if ncessary
                adWin.RibbonControl ribbon = adWin.ComponentManager.Ribbon;
                adWin.RibbonTab tab = null;
                bool defaultTab = false;

                foreach (adWin.RibbonTab t in ribbon.Tabs)
                {
                    if (t.Id == tabName)
                    {
                        if (t.Id != t.Name)
                            defaultTab = true;
                        tab = t;
                        break;
                    }
                }

                if (!defaultTab && tab == null)
                    revApp.CreateRibbonTab(tabName);
                if (defaultTab)
                    tab = null;

                // Verify if the panel exists
                List<RibbonPanel> panels;
                if (defaultTab)
                    panels = revApp.GetRibbonPanels();
                else
                    panels = revApp.GetRibbonPanels(tabName);

                RibbonPanel panel = null;
                foreach (RibbonPanel rp in panels)
                {
                    if (rp.Name == panelName)
                    {
                        panel = rp;
                        break;
                    }
                }

                if (panel == null && !defaultTab)
                    panel = revApp.CreateRibbonPanel(tabName, panelName);
                else if (panel == null && defaultTab)
                    panel = revApp.CreateRibbonPanel(panelName);

                return panel;
            }
            catch
            {
                return null;
            }
        }
    }
}
