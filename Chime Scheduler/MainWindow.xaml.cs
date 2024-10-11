using System.Windows;
using System.Diagnostics;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using Application = System.Windows.Application;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog; // Add this for Icon

namespace ChimeScheduler
{
    public partial class MainWindow : Window
    {
        private string? scriptPath;
        private int intervalMinutes;
        private CancellationTokenSource? cancellationTokenSource;
        private bool isRunning = false;
        private DateTime? startTime;
        private NotifyIcon? trayIcon;

        public MainWindow()
        {
            InitializeComponent();
            DpStartDate.SelectedDate = DateTime.Today;
            TxtStartTime.Text = DateTime.Now.ToString("HH:mm");
            InitializeTrayIcon();
        }

        private void InitializeTrayIcon()
        {

            trayIcon = new NotifyIcon
            {
                Icon = LoadIconFromResource("favicon.ico"),
                Visible = false,
                ContextMenuStrip = new ContextMenuStrip()
            };

            trayIcon.ContextMenuStrip.Items.Add("Show", null, ShowMenuItem_Click!);
            trayIcon.ContextMenuStrip.Items.Add("Exit", null, ExitMenuItem_Click!);

            trayIcon.DoubleClick += TrayIcon_DoubleClick!;

            StateChanged += MainWindow_StateChanged!;
        }

        private Icon LoadIconFromResource(string resourceName)
        {
            // Load the icon from the resources
            using (Stream stream = Application.GetResourceStream(new Uri($"pack://application:,,,/{resourceName}")).Stream)
            {
                return new Icon(stream); // Create an Icon from the stream
            }
        }

        private void MainWindow_StateChanged(object sender, System.EventArgs e)
        {
            if (WindowState == WindowState.Minimized)
            {
                Hide();
                trayIcon.Visible = true;
            }
        }

        private void TrayIcon_DoubleClick(object sender, EventArgs e)
        {
            Show();
            WindowState = WindowState.Normal;
            trayIcon.Visible = false;
        }

        private void ShowMenuItem_Click(object sender, EventArgs e)
        {
            Show();
            WindowState = WindowState.Normal;
            trayIcon.Visible = false;
        }

        private void ExitMenuItem_Click(object sender, EventArgs e)
        {
            trayIcon.Visible = false;
            trayIcon.Dispose();
            Application.Current.Shutdown();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            trayIcon.Visible = false;
            trayIcon.Dispose();
        }

        private void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "PowerShell Scripts (*.ps1)|*.ps1"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                scriptPath = openFileDialog.FileName;
                TxtScriptPath.Text = scriptPath;
            }
        }

        private void ChkUseStartTime_Checked(object sender, RoutedEventArgs e)
        {

            TxtStartTime.IsEnabled = true;
            DpStartDate.IsEnabled = true;
        }

        private void ChkUseStartTime_Unchecked(object sender, RoutedEventArgs e)
        {
            TxtStartTime.IsEnabled = false;
            DpStartDate.IsEnabled = false;
        }

        private async void BtnStartStop_Click(object sender, RoutedEventArgs e)
        {
            if (!isRunning)
            {
                if (string.IsNullOrEmpty(scriptPath) || !File.Exists(scriptPath))
                {
                    MessageBox.Show("Please select a valid PowerShell script.");
                    return;
                }

                if (!int.TryParse(TxtInterval.Text, out intervalMinutes) || intervalMinutes <= 0)
                {
                    MessageBox.Show("Please enter a valid interval in minutes.");
                    return;
                }

                if (ChkUseStartTime.IsChecked == true)
                {
                    if (!TryParseStartTime(out DateTime parsedStartTime))
                    {
                        MessageBox.Show("Please enter a valid start time.");
                        return;
                    }
                    startTime = parsedStartTime;
                }
                else
                {
                    startTime = null;
                }

                isRunning = true;
                BtnStartStop.Content = "Stop";
                cancellationTokenSource = new CancellationTokenSource();
                try
                {
                    await RunScriptPeriodicallyAsync(cancellationTokenSource.Token);
                }
                catch (OperationCanceledException)
                {
                    // This exception is expected when cancelling the task
                }
                finally
                {
                    isRunning = false;
                    BtnStartStop.Content = "Start";
                }
            }
            else
            {
                cancellationTokenSource.Cancel();
            }
        }

        private bool TryParseStartTime(out DateTime result)
        {
            if (DpStartDate.SelectedDate.HasValue && DateTime.TryParse(TxtStartTime.Text, out DateTime parsedTime))
            {
                result = DpStartDate.SelectedDate.Value.Date + parsedTime.TimeOfDay;
                return true;
            }
            result = DateTime.MinValue;
            return false;
        }

        private async Task RunScriptPeriodicallyAsync(CancellationToken cancellationToken)
        {
            if (startTime.HasValue)
            {
                TimeSpan delay = startTime.Value - DateTime.Now;
                if (delay > TimeSpan.Zero)
                {
                    await Task.Delay(delay, cancellationToken);
                }
            }

            while (!cancellationToken.IsCancellationRequested)
            {
                await RunPowerShellScriptAsync(scriptPath, cancellationToken);
                await Task.Delay(TimeSpan.FromMinutes(intervalMinutes), cancellationToken);
            }
        }

        private async Task RunPowerShellScriptAsync(string scriptPath, CancellationToken cancellationToken)
        {
            using (Process process = new Process())
            {
                process.StartInfo.FileName = "powershell.exe";
                process.StartInfo.Arguments = $"-ExecutionPolicy Bypass -File \"{scriptPath}\"";
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.CreateNoWindow = true;

                process.Start();
                string output = await process.StandardOutput.ReadToEndAsync();
                await process.WaitForExitAsync(cancellationToken);

                Dispatcher.Invoke(() =>
                {
                    TxtOutput.Text += $"{DateTime.Now}: Script executed\n{output}\n\n";
                    TxtOutput.ScrollToEnd();
                });
            }
        }
    }
}