using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace RevitCommon
{
    /// <summary>
    /// Interaction logic for UILocation.xaml
    /// </summary>
    public partial class UILocation : Window
    {
        LinearGradientBrush eBrush = new LinearGradientBrush(Color.FromArgb(255, 195, 195, 195), Color.FromArgb(255, 245, 245, 245), new Point(0, 0), new Point(0, 1));
        SolidColorBrush lBrush = new SolidColorBrush(Color.FromArgb(0, 0, 0, 0));

        string tabName;
        string panelName;

        public string Tab { get { return tabName; } }
        public string Panel { get { return panelName; } }

        /// <summary>
        /// Form to allow changing of the location of the button in the Ribbon UI
        /// </summary>
        /// <param name="command">Name of the command. Will appear in the title.</param>
        /// <param name="tab">Name of the tab it currently resides</param>
        /// <param name="panel">Name of the panel it currently resides</param>
        public UILocation(string command, string tab, string panel)
        {
            tabName = tab;
            panelName = panel;
            InitializeComponent();
            titleLabel.Content = "UI - " + command;

        
            tabTextBox.Text = tabName;
            panelTextBox.Text = panelName;
        }

        private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void cancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void cancelButton_MouseEnter(object sender, MouseEventArgs e)
        {
            cancelRect.Fill = eBrush;
        }

        private void cancelButton_MouseLeave(object sender, MouseEventArgs e)
        {
            cancelRect.Fill = lBrush;
        }

        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            if (tabTextBox.Text != null && tabTextBox.Text != string.Empty)
                tabName = tabTextBox.Text;
            if (panelTextBox.Text != null && panelTextBox.Text != string.Empty)
                panelName = panelTextBox.Text;

            Close();
        }

        private void okButton_MouseEnter(object sender, MouseEventArgs e)
        {
            okButtonRect.Fill = eBrush;
        }

        private void okButton_MouseLeave(object sender, MouseEventArgs e)
        {
            okButtonRect.Fill = lBrush;
        }
    }
}
