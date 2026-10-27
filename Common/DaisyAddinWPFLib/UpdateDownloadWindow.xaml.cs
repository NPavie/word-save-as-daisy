using System;
using System.ComponentModel;
using System.IO;
using System.Threading;
using System.Windows;

namespace Daisy.SaveAsDAISY.WPF
{
    /// <summary>
    /// Dialog showing the download progress of an update installer,
    /// with a cancel button. The download is canceled when the window is closed.
    /// </summary>
    public partial class UpdateDownloadWindow : Window
    {
        private readonly Daisy.SaveAsDAISY.Conversion.Update.UpdateInfo update;
        private readonly string destinationPath;
        private readonly CancellationTokenSource cancellation = new CancellationTokenSource();
        private bool downloadCompleted = false;

        public bool DownloadSucceeded { get; private set; }
        public string DownloadedInstallerPath { get; private set; }
        public Exception DownloadError { get; private set; }

        public UpdateDownloadWindow(Daisy.SaveAsDAISY.Conversion.Update.UpdateInfo update, string destinationFolder)
        {
            InitializeComponent();
            this.update = update;
            string fileName = string.IsNullOrEmpty(update.InstallerFileName)
                ? "SaveAsDAISYInstaller.exe"
                : update.InstallerFileName;
            destinationPath = System.IO.Path.Combine(destinationFolder, fileName);
            UpdateText.Text = "Downloading Save as DAISY update " + update.Version
                + (update.IsPrerelease ? " (beta)" : "")
                + ":\r\n" + fileName;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            ThreadPool.QueueUserWorkItem(delegate
            {
                try
                {
                    Daisy.SaveAsDAISY.Conversion.Update.UpdateManager.DownloadUpdate(
                        update,
                        destinationPath,
                        delegate (long received, long total, int percent)
                        {
                            Dispatcher.BeginInvoke(new Action(delegate
                            {
                                DownloadProgress.Value = percent;
                                ProgressText.Text = total > 0
                                    ? FormatBytes(received) + " / " + FormatBytes(total) + " (" + percent + "%)"
                                    : FormatBytes(received);
                            }));
                        },
                        cancellation.Token
                    );
                    downloadCompleted = true;
                    DownloadSucceeded = true;
                    DownloadedInstallerPath = destinationPath;
                    Dispatcher.BeginInvoke(new Action(delegate
                    {
                        DialogResult = true;
                        Close();
                    }));
                }
                catch (OperationCanceledException)
                {
                    downloadCompleted = true;
                    Dispatcher.BeginInvoke(new Action(delegate
                    {
                        DialogResult = false;
                        Close();
                    }));
                }
                catch (Exception ex)
                {
                    downloadCompleted = true;
                    DownloadError = ex;
                    Dispatcher.BeginInvoke(new Action(delegate
                    {
                        MessageBox.Show(
                            "The update download failed:\r\n" + ex.Message,
                            "Save as DAISY update",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error
                        );
                        DialogResult = false;
                        Close();
                    }));
                }
            });
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            cancellation.Cancel();
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            if (!downloadCompleted)
            {
                cancellation.Cancel();
            }
        }

        private static string FormatBytes(long bytes)
        {
            string[] units = { "B", "KB", "MB", "GB" };
            double value = bytes;
            int unit = 0;
            while (value >= 1024 && unit < units.Length - 1)
            {
                value /= 1024;
                unit++;
            }
            return string.Format("{0:0.#} {1}", value, units[unit]);
        }
    }
}
