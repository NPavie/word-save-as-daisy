using Daisy.SaveAsDAISY.Conversion;
using System;
using System.IO;
using System.Windows;

namespace Daisy.SaveAsDAISY.WPF
{
    /// <summary>
    /// Front-end of the addin auto-update process: check for updates on the GitHub
    /// releases, ask the user, download with progress feedback, then run the installer.
    /// </summary>
    public class Updater
    {
        public const string InstallerDownloadFolder = "updates";

        /// <summary>
        /// Check for an update according to the user settings channel, and if one is
        /// available, ask the user whether to install it. Runs the dialogs on a
        /// dedicated thread (same pattern as the conversion progress dialog).
        /// </summary>
        public static void CheckForUpdateAndPrompt()
        {
            ConverterSettings settings = ConverterSettings.Instance;
            Daisy.SaveAsDAISY.Conversion.Update.UpdateInfo update;
            try
            {
                update = Daisy.SaveAsDAISY.Conversion.Update.UpdateManager.CheckForUpdate(settings.UpdateChannel);
            }
            catch (Exception e)
            {
                AddinLogger.Error("Error while checking for Save as DAISY updates", e);
                return;
            }
            if (update == null)
            {
                return;
            }
            RunOnStaThread(delegate
            {
                MessageBoxResult result = MessageBox.Show(
                    "A new version of Save as DAISY is available: " + update.Version
                    + (update.IsPrerelease ? " (pre-release)" : "")
                    + "\r\n\r\nDo you want to download and install this update?\r\n"
                    + "(Microsoft Word will be closed by the installer)",
                    "Save as DAISY update",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question
                );
                if (result == MessageBoxResult.Yes)
                {
                    DownloadAndRunInstaller(update);
                }
            });
        }

        /// <summary>
        /// Show the download progress dialog, then launch the downloaded installer.
        /// </summary>
        public static void DownloadAndRunInstaller(Daisy.SaveAsDAISY.Conversion.Update.UpdateInfo update)
        {
            string destinationFolder = Path.Combine(ConverterSettings.ApplicationDataFolder, InstallerDownloadFolder);
            try
            {
                if (!Directory.Exists(destinationFolder))
                {
                    Directory.CreateDirectory(destinationFolder);
                }
            }
            catch (Exception e)
            {
                AddinLogger.Error("Could not create the update download folder", e);
                MessageBox.Show(
                    "The update download folder could not be created:\r\n" + e.Message,
                    "Save as DAISY update",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );
                return;
            }
            UpdateDownloadWindow downloadWindow = new UpdateDownloadWindow(update, destinationFolder);
            bool? downloaded = downloadWindow.ShowDialog();
            if (downloaded == true)
            {
                try
                {
                    Daisy.SaveAsDAISY.Conversion.Update.UpdateManager.RunInstaller(downloadWindow.DownloadedInstallerPath);
                }
                catch (Exception e)
                {
                    AddinLogger.Error("Could not start the update installer", e);
                    MessageBox.Show(
                        "The update installer could not be started:\r\n" + e.Message,
                        "Save as DAISY update",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error
                    );
                }
            }
        }

        private static void RunOnStaThread(Action action)
        {
            System.Threading.Thread thread = new System.Threading.Thread(new System.Threading.ThreadStart(action));
            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
        }
    }
}
