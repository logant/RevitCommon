using System.Collections.Generic;
using Autodesk.Revit.UI;
using adWin = Autodesk.Windows;
using System;
using System.Runtime.InteropServices;

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
        /// <returns>True if successful, False if it failed to add a button to the Ribbon.</returns>
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
        /// This method is similar to the previous one in that it creates a set of buttons on a specified tab and panel, but this one
        /// adds a separator to the end of the buttons. 
        /// </summary>
        /// <param name="revApp">UIControlledApplication from an IExternalApplication StartUp method</param>
        /// <param name="tabName">The name of the tab that you want to add a button to. This must either be 'Add-Ins' or a non-standard tab.</param>
        /// <param name="panelName">The name of the panel that you want to add a button to.</param>
        /// <param name="buttons">One or more buttons to add to the panel.</param>
        /// <param name="addSeparators">Add a separator AFTER the buttons are created. This should only be True when you know additional buttons will be added</param>
        /// <returns>True if successful, False if it failed to add a button to the Ribbon.</returns>
        public static bool AddToRibbon(UIControlledApplication revApp, string tabName, string panelName, List<PushButtonData> buttons, bool addSeparators)
        {
            RibbonPanel panel = GetRibbonPanel(revApp, tabName, panelName, out bool existingPanel);

            // Add the button(s) to the panel
            if (panel != null)
            {
                if (addSeparators && existingPanel)
                    panel.AddSeparator();
                foreach (PushButtonData pbd in buttons)
                {
                    panel.AddItem(pbd);
                }
                if (addSeparators)
                    panel.AddSeparator();
                return true;
            }
            else
            {
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                foreach (PushButtonData pbd in buttons)
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
        /// <returns>True if successful, False if it failed to add a button to the Ribbon.</returns>
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
        /// <param name="button">SplitButtonData object to add to the ribbon.</param>
        /// <returns>If successful, a SplitButton object is returned that can be used to add commands from its drop-down. If unsuccessful, it returns null.</returns>
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
        /// Add a PullDownbutton to Revit. Similar to the SplitButton created in the above method, but PullDownButtons
        /// do not have a default command and will only show the list of command options when clicked.
        /// </summary>
        /// <param name="revApp">Revit's UIControlledApplication for adding the button</param>
        /// <param name="tabName">Name of the tab you want to add the button to.</param>
        /// <param name="panelName">Name of the panel on the tab you want to add the button</param>
        /// <param name="button">PullDownButtonData to add as a PullDownButton</param>
        /// <returns>A PulldownButton object that you can add PushDownButtonData objects to.</returns>
        public static PulldownButton AddToRibbon(UIControlledApplication revApp, string tabName, string panelName, PulldownButtonData button)
        {
            RibbonPanel panel = GetRibbonPanel(revApp, tabName, panelName);

            // Add the button to the panel
            if (panel != null)
                return panel.AddItem(button) as PulldownButton;
            else
            {
                TaskDialog.Show("Error", "Could not add pull down button to the Revit ribbon for:\n" + button.Text);
                return null;
            }
        }

        /// <summary>
        /// Create a stack of 3 small buttons in the Revit UI.
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

        /// <summary>
        /// Create a stack of 3 buttons in the Revit UI.
        /// </summary>
        /// <param name="revApp">Revit's UIControlledApplication for adding the button</param>
        /// <param name="tabName">Name of the tab you want to add the button to.</param>
        /// <param name="panelName">Name of the panel on the tab you want to add the button</param>
        /// <param name="button0">Top button in the stack</param>
        /// <param name="button1">Middle button in the stack</param>
        /// <param name="button2">Bottom button in the stack</param>
        /// <param name="stackName">Name of the stacked panel that these buttons belong to.</param>
        public static bool AddToRibbon(UIControlledApplication revApp, string tabName, string panelName, PushButtonData button0, PushButtonData button1, PushButtonData button2, string stackName)
        {
            RibbonPanel panel = GetRibbonPanel(revApp, tabName, panelName);

            // Add the button to the panel
            if (panel != null)
                panel.AddStackedItems(button0, button1, button2);
            else
                TaskDialog.Show("Error", "Could not create stacked buttons for:\n" + button0.Text + "\n" + button1.Text + "\n" + button2.Text);

            adWin.RibbonItem i = null;
            foreach(adWin.RibbonTab t in adWin.ComponentManager.Ribbon.Tabs)
            {
                foreach(adWin.RibbonPanel p in t.Panels)
                {
                    if(p.Source.Name == panelName)
                    {
                        foreach(adWin.RibbonItem item in p.Source.Items)
                        {
                            i = item;
                        }
                    }
                }
            }
            // Try to set the name.
            if (i is adWin.RibbonRowPanel) // Stacked buttons
                i.Text = stackName;

            return true;
        }

        /// <summary>
        /// This is used by the other methods in this class, it's purpose is to find or create the tab and panel specified by
        /// the inputs. If the items do not exist they get created, if they do exist they're found and returned. This should only
        /// be used with a tab name that is non-default to the Revit Ribbon, excepting the Add-Ins tab which is allowed.
        /// </summary>
        /// <param name="revApp">UIControlledApplication from the IExternalApplication's OnStartUp method.</param>
        /// <param name="tabName">Name of the tab a button will be created on. Only Add-Ins is acceptable from the default Revit tabs.</param>
        /// <param name="panelName">Name of the panel the button will be created on.</param>
        /// <returns></returns>
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

        private static RibbonPanel GetRibbonPanel(UIControlledApplication revApp, string tabName, string panelName, out bool panelExists)
        {
            try
            {
                // Verify if the tab exists, create it if necessary
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
                panelExists = false;
                foreach (RibbonPanel rp in panels)
                {
                    if (rp.Name == panelName)
                    {
                        panel = rp;
                        panelExists = true;
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
                panelExists = false;
                return null;
            }
        }
    }
}
