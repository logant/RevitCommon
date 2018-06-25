using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace RevitCommon
{
    /// <summary>
    /// Interaction logic for NotificationWindow.xaml
    /// </summary>
    public partial class NotificationWindow : Window
    {
        DoubleAnimation animation;
        int delay = 4;

        /// <summary>
        /// Launches a temporary notification window with a specified image as the background
        /// The text overlay will be a stroked text block with white stroke, red fill.
        /// </summary>
        /// <param name="title">Shows up in the title bar, top left of form.</param>
        /// <param name="message">Text block that shows up in the main message of the form.</param>
        /// <param name="filePath">Specified file path to use for the form's background</param>
        public NotificationWindow(string title, string message, string filePath)
        {
            InitializeComponent();

            if (System.IO.File.Exists(filePath))
            {
                ImageSource imgSrc = new BitmapImage(new Uri(filePath));
                imageBorder.Background = new ImageBrush(imgSrc);
            }
            titleLabel.Content = title;
            msgTextBlock.Text = message;

            DispatcherTimer timer = new DispatcherTimer();
            timer.Interval = TimeSpan.FromSeconds(delay);
            timer.Tick += timer_Tick;
            timer.Start();

        }

        /// <summary>
        /// Launches a temporary notification window with an optional image background
        /// The image underlay comes from a curated set of images in tlogan@hksinc.com's 
        /// LINE folder. When an image is used, the text appears as a red fill with white outline
        /// in a bold font weight, while without an image it is regular font weight and black
        /// with no stroke outline.
        /// </summary>
        /// <param name="title">Shows up in the title bar, top left of the form.</param>
        /// <param name="message">Text block that shows up in the main message of the form.</param>
        /// <param name="randomImg">bool choice on whether to search for an image or not.</param>
        public NotificationWindow(string title, string message, bool randomImg)
        {
            InitializeComponent();

            if (randomImg)
            {
                string imgDir = HKS.GetPath("ee-image-path");
                if (string.IsNullOrEmpty(imgDir))
                    return;
                
                string filePath = string.Empty;
                if(System.IO.Directory.Exists(imgDir))
                {
                    List<System.IO.FileInfo> imageFiles = new System.IO.DirectoryInfo(imgDir).GetFiles("*.jpg").ToList();
                    imageFiles.AddRange(new System.IO.DirectoryInfo(imgDir).GetFiles("*.png"));
                    if(imageFiles.Count > 0)
                    {
                        Random r = new Random();
                        filePath = imageFiles[r.Next(imageFiles.Count - 1)].FullName;
                    }
                }
                if (System.IO.File.Exists(filePath))
                {
                    ImageSource imgSrc = new BitmapImage(new Uri(filePath));
                    imageBorder.Background = new ImageBrush(imgSrc);
                }
            }
            else
            {
                // Modify the stroke textblock to simplify it.
                msgTextBlock.StrokeThickness = 0;
                msgTextBlock.FontWeight = FontWeights.Normal;
                msgTextBlock.Fill = new SolidColorBrush(Colors.Black);
            }
            titleLabel.Content = title;
            msgTextBlock.Text = message;

            DispatcherTimer timer = new DispatcherTimer();
            timer.Interval = TimeSpan.FromSeconds(delay);
            timer.Tick += timer_Tick;
            timer.Start();

        }


        void timer_Tick(object sender, EventArgs e)
        {
            DispatcherTimer t = (DispatcherTimer)sender;
            t.Stop();
            t.Tick -= timer_Tick;


            animation = new DoubleAnimation(0, (Duration)TimeSpan.FromSeconds(3));
            animation.Completed += (s, _) => Close();
            this.BeginAnimation(UIElement.OpacityProperty, animation);
        }

        private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                DragMove();
            }
            catch { }
        }

        private void closeButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Kill the close animation
                this.BeginAnimation(UIElement.OpacityProperty, null);
                animation.Completed -= (s, _) => Close();
            }
            catch { }
            // Close the form
            Close();
        }

        private void Border_MouseEnter(object sender, MouseEventArgs e)
        {
            try
            {
                // Kill the close animation
                BeginAnimation(OpacityProperty, null);
                animation.Completed -= (s, _) => Close();
                Opacity = 1.0;
            }
            catch { } // Animation or something is null. Just let it finishing closing the form
        }

        private void Border_MouseLeave(object sender, MouseEventArgs e)
        {
            try
            {
                // Restart the close animation
                animation = new DoubleAnimation(0, (Duration)TimeSpan.FromSeconds(3));
                animation.Completed += (s, _) => Close();
                BeginAnimation(UIElement.OpacityProperty, animation);
            }
            catch { }
        }
    }
}
