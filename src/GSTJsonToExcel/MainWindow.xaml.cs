using System;
using System.Windows;
using System.Windows.Input;
using GSTJsonToExcel.ViewModels;

namespace GSTJsonToExcel
{
    public partial class MainWindow : Window
    {
        private readonly MainViewModel _viewModel;

        public MainWindow(MainViewModel viewModel)
        {
            InitializeComponent();
            _viewModel = viewModel;
            DataContext = _viewModel;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                var workArea = SystemParameters.WorkArea;
                if (Height > workArea.Height)
                {
                    Height = Math.Max(MinHeight, workArea.Height - 30);
                }
                if (Width > workArea.Width)
                {
                    Width = Math.Max(MinWidth, workArea.Width - 30);
                }
                if (Top < workArea.Top)
                {
                    Top = workArea.Top;
                }
                if (Left < workArea.Left)
                {
                    Left = workArea.Left;
                }
            }
            catch
            {
                // Fallback gracefully
            }
        }

        private void Header_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                MaximizeButton_Click(sender, e);
            }
            else
            {
                try
                {
                    DragMove();
                }
                catch
                {
                    // Ignore if mouse state changed
                }
            }
        }

        private void MinimizeButton_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        private void MaximizeButton_Click(object sender, RoutedEventArgs e)
        {
            if (WindowState == WindowState.Maximized)
            {
                WindowState = WindowState.Normal;
                if (MaximizeButton != null) MaximizeButton.Content = "🗖";
            }
            else
            {
                WindowState = WindowState.Maximized;
                if (MaximizeButton != null) MaximizeButton.Content = "🗗";
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Window_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        private async void Window_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[]? droppedPaths = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (droppedPaths != null && droppedPaths.Length > 0)
                {
                    await _viewModel.LoadPathsAsync(droppedPaths);
                }
            }
        }
    }
}