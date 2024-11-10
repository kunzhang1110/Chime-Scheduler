using System.Windows;
using System.Diagnostics;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using System.Text.Json;
using Application = System.Windows.Application;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using System.Windows.Controls;
using Microsoft.Win32;

namespace ScriptScheduler
{
    public partial class MainWindow : Window
    {
        private readonly string configFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            System.Reflection.Assembly.GetExecutingAssembly().GetName().Name ?? "ScriptScheduler",
            "configs.json"
        );//C:\Users\{YourUsername}\AppData\Roaming

        private string? scriptPath;
        private int intervalMinutes;
        private DateTime? startTime;
        private NotifyIcon? trayIcon;
        private CancellationTokenSource? cancellationTokenSource;
        private bool isRunning = false;
        private Dictionary<string, string>? configs;

        public MainWindow()
        {
            InitializeComponent();
            InitializeTrayIcon();
            DpStartDate.SelectedDate = DateTime.Today;
            TxtStartTime.Text = DateTime.Now.ToString("HH:mm");
            Loaded += MainWindow_Loaded;
            Closing += MainWindow_Closing!;
            //SystemEvents.PowerModeChanged += SystemEvents_PowerModeChanged;
            SystemEvents.SessionSwitch += SystemEvents_HandleSessionSwitch;
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

        private static Icon LoadIconFromResource(string resourceName)
        {
            using Stream stream =
                Application.GetResourceStream(new Uri($"pack://application:,,,/{resourceName}")).Stream;
            return new Icon(stream);
        }

        private void MainWindow_StateChanged(object sender, EventArgs e)
        {
            if (WindowState == WindowState.Minimized)
            {
                Hide();
                trayIcon!.Visible = true;
            }
        }

        private void TrayIcon_DoubleClick(object sender, EventArgs e)
        {
            Show();
            WindowState = WindowState.Normal;
            trayIcon!.Visible = false;
        }

        private void ShowMenuItem_Click(object sender, EventArgs e)
        {
            Show();
            WindowState = WindowState.Normal;
            trayIcon!.Visible = false;
        }

        private void ExitMenuItem_Click(object sender, EventArgs e)
        {
            trayIcon!.Visible = false;
            trayIcon.Dispose();
            Application.Current.Shutdown();
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            if (File.Exists(configFilePath))
            {
                string json = File.ReadAllText(configFilePath);
                configs = JsonSerializer.Deserialize<Dictionary<string, string>>(json)
                    ?? new Dictionary<string, string>();
                if (configs.TryGetValue("scriptPath", out var savedScriptPath))
                {
                    scriptPath = savedScriptPath;
                    TxtScriptPath.Text = scriptPath;
                }
                if (configs.TryGetValue("intervalMinutes", out var savedIntervalMinutes))
                {
                    TxtInterval.Text = savedIntervalMinutes;
                }
            }
            else
            {
                configs = new Dictionary<string, string>();
            }
        }

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {   // Save the configs and dispose tray icon when closing the window
            configs!["scriptPath"] = scriptPath ?? string.Empty;
            configs!["intervalMinutes"] = intervalMinutes.ToString();

            Directory.CreateDirectory(Path.GetDirectoryName(configFilePath)!);
            string json = JsonSerializer.Serialize(configs, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(configFilePath, json);

            trayIcon!.Visible = false;
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

        private void ChkUseStartTime_CheckedUnchecked(object sender, RoutedEventArgs e)
        {
            bool isChecked = ChkUseStartTime.IsChecked == true; // Check the state of the CheckBox

            TxtStartTime.IsEnabled = isChecked; // Set the TextBox enabled state
            DpStartDate.IsEnabled = isChecked;   // Set the DatePicker enabled state
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
                cancellationTokenSource = new CancellationTokenSource();
                BtnStartStop.Content = "Stop";
                SetControlIsEnabledState(false);

                try
                {
                    await RunScriptPeriodicallyAsync(cancellationTokenSource.Token);
                }
                catch (OperationCanceledException)
                {
                    Console.WriteLine("Operation cancelled.");
                }
                finally
                {
                    isRunning = false;
                    BtnStartStop.Content = "Start";
                    SetControlIsEnabledState(true);
                }
            }
            else
            {
                cancellationTokenSource?.Cancel(); //request to cancel
            }

        }

        private void SetControlIsEnabledState(bool enable)
        {
            TxtScriptPath.IsEnabled = enable;
            BtnBrowse.IsEnabled = enable;
            TxtInterval.IsEnabled = enable;
            ChkUseStartTime.IsEnabled = enable;
            TxtStartTime.IsEnabled = enable;
            DpStartDate.IsEnabled = enable;
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
                TimeSpan elapsedTime = DateTime.Now - startTime.Value;

                if (elapsedTime.TotalSeconds > 0)
                {
                    // If the start time is in the past
                    double intervalsPassed = Math.Floor(elapsedTime.TotalMinutes / intervalMinutes);
                    TimeSpan timeUntilNextExecution = TimeSpan.FromMinutes((intervalsPassed + 1) * intervalMinutes) - elapsedTime;

                    await Task.Delay(timeUntilNextExecution, cancellationToken);
                }
                else
                {
                    // If the start time is in the future, wait until the start time
                    await Task.Delay(-elapsedTime, cancellationToken);
                }
            }

            while (!cancellationToken.IsCancellationRequested)
            {
                await RunPowerShellScriptAsync(scriptPath!, cancellationToken);
                await Task.Delay(TimeSpan.FromMinutes(intervalMinutes), cancellationToken);
            }
        }

        private async Task RunPowerShellScriptAsync(string scriptPath, CancellationToken cancellationToken)
        {
            using Process process = new();
            process.StartInfo.FileName = "powershell.exe";
            process.StartInfo.Arguments = $"-ExecutionPolicy Bypass -File \"{scriptPath}\"";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.CreateNoWindow = true;
            process.Start();

            string output = await process.StandardOutput.ReadToEndAsync();
            await process.WaitForExitAsync(cancellationToken); //wait for the powershell process to exit

            Dispatcher.Invoke(() =>
            {   //update UI thread from background thread with Dispatcher
                TxtOutput.Text += $"{DateTime.Now}: Script executed\n{output}\n\n";
                TxtOutput.ScrollToEnd();
            });
        }

        private async void SystemEvents_HandleSessionSwitch(object sender, SessionSwitchEventArgs e)
        {
      
            if (e.Reason == SessionSwitchReason.SessionUnlock)
            {
                cancellationTokenSource!.Cancel();
                cancellationTokenSource!.Dispose();
                await Task.Delay(100);

                isRunning = true;
                cancellationTokenSource = new CancellationTokenSource();
                BtnStartStop.Content = "Stop";
                SetControlIsEnabledState(false);

                try
                {
                    await RunScriptPeriodicallyAsync(cancellationTokenSource.Token);
                }
                catch (OperationCanceledException)
                {
                    Console.WriteLine("Operation cancelled.");

                }
                finally
                {
                    isRunning = false;
                    BtnStartStop.Content = "Start";
                    SetControlIsEnabledState(true);
                }
            }
        }
    }
}