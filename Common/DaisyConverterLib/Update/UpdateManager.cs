using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Threading;

namespace Daisy.SaveAsDAISY.Conversion.Update
{
    /// <summary>
    /// Description of an available update, extracted from a GitHub release
    /// </summary>
    public class UpdateInfo
    {
        public Version Version { get; set; }
        public string TagName { get; set; }
        public string Name { get; set; }
        public string NotesUrl { get; set; }
        public bool IsPrerelease { get; set; }
        public string InstallerUrl { get; set; }
        public string InstallerFileName { get; set; }
        public long InstallerSize { get; set; }
    }

    /// <summary>
    /// Checks for updates on the project GitHub releases and downloads the installer.<br/>
    /// Releases are expected to embed an installer executable asset (bootstrapping an MSI),
    /// similar to the existing "SaveAsDAISYInstaller-x.y.z.w.exe" release assets.
    /// </summary>
    public class UpdateManager
    {
        public const string ReleasesApiUrl = "https://api.github.com/repos/NPavie/word-save-as-daisy/releases";
        public const string ReleasesPageUrl = "https://github.com/NPavie/word-save-as-daisy/releases";
        private const string UserAgent = "SaveAsDAISY-WordAddIn-Updater";

        /// <summary>
        /// Version of the currently running assembly
        /// </summary>
        public static Version CurrentVersion
        {
            get { return Assembly.GetExecutingAssembly().GetName().Version; }
        }

        /// <summary>
        /// Query the GitHub releases and return the latest applicable update for the given channel,
        /// or null if no update is available.
        /// </summary>
        /// <param name="channel">The update channel (stable = only non pre-releases, beta = also pre-releases)</param>
        /// <param name="timeoutMilliseconds">Timeout for the API request</param>
        public static UpdateInfo CheckForUpdate(UpdateChannel.Enum channel, int timeoutMilliseconds = 15000)
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            List<UpdateInfo> releases = FetchReleases(timeoutMilliseconds);
            UpdateInfo candidate = null;
            Version current = CurrentVersion;
            foreach (UpdateInfo release in releases)
            {
                if (release.Version == null || release.InstallerUrl == null)
                {
                    continue;
                }
                if (channel == UpdateChannel.Enum.Stable && release.IsPrerelease)
                {
                    continue;
                }
                if (release.Version > current && (candidate == null || release.Version > candidate.Version))
                {
                    candidate = release;
                }
            }
            return candidate;
        }

        /// <summary>
        /// Download the update installer to the given destination path, reporting progress through the callback.
        /// </summary>
        /// <param name="update">The update to download</param>
        /// <param name="destinationPath">Full path of the resulting installer file</param>
        /// <param name="progressCallback">Called with (bytesReceived, totalBytes, percent), on a worker thread</param>
        /// <param name="cancellationToken">Token used to cancel the download</param>
        public static void DownloadUpdate(UpdateInfo update, string destinationPath,
            Action<long, long, int> progressCallback, CancellationToken cancellationToken)
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            using (HttpClient client = new HttpClient())
            {
                client.Timeout = TimeSpan.FromMinutes(30);
                using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, update.InstallerUrl))
                using (HttpResponseMessage response = client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken).GetAwaiter().GetResult())
                {
                    response.EnsureSuccessStatusCode();
                    long total = update.InstallerSize > 0
                        ? update.InstallerSize
                        : (response.Content.Headers.ContentLength ?? -1);
                    string tempPath = destinationPath + ".part";
                    try
                    {
                        using (Stream downloadStream = response.Content.ReadAsStreamAsync().GetAwaiter().GetResult())
                        using (FileStream fileStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None))
                        {
                            byte[] buffer = new byte[81920];
                            long received = 0;
                            int read;
                            while ((read = downloadStream.Read(buffer, 0, buffer.Length)) > 0)
                            {
                                cancellationToken.ThrowIfCancellationRequested();
                                fileStream.Write(buffer, 0, read);
                                received += read;
                                if (progressCallback != null)
                                {
                                    int percent = total > 0 ? (int)(received * 100 / total) : 0;
                                    progressCallback(received, total, percent);
                                }
                            }
                        }
                        if (File.Exists(destinationPath))
                        {
                            File.Delete(destinationPath);
                        }
                        File.Move(tempPath, destinationPath);
                    }
                    finally
                    {
                        try
                        {
                            if (File.Exists(tempPath))
                            {
                                File.Delete(tempPath);
                            }
                        }
                        catch (Exception) { }
                    }
                }
            }
        }

        /// <summary>
        /// Launch the downloaded update installer. The add-in is expected to be closed by the
        /// installer deployment (Word should be closed by the user when asked).
        /// </summary>
        /// <param name="installerPath">Full path of the installer executable</param>
        public static void RunInstaller(string installerPath)
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = installerPath,
                UseShellExecute = true
            });
        }

        private static List<UpdateInfo> FetchReleases(int timeoutMilliseconds)
        {
            using (HttpClient client = new HttpClient())
            {
                client.Timeout = TimeSpan.FromMilliseconds(timeoutMilliseconds);
                client.DefaultRequestHeaders.Add("User-Agent", UserAgent);
                client.DefaultRequestHeaders.Add("Accept", "application/vnd.github+json");
                using (HttpResponseMessage response = client.GetAsync(ReleasesApiUrl).GetAwaiter().GetResult())
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new HttpRequestException(
                            "GitHub releases API returned " + (int)response.StatusCode + " " + response.ReasonPhrase);
                    }
                    string json = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                    return ParseReleases(json);
                }
            }
        }

        private static List<UpdateInfo> ParseReleases(string json)
        {
            List<UpdateInfo> releases = new List<UpdateInfo>();
            // Minimal JSON parsing of the releases array (no JSON dependency in this project)
            Newtonsoft.Json.Linq.JArray array = Newtonsoft.Json.Linq.JArray.Parse(json);
            foreach (Newtonsoft.Json.Linq.JObject release in array.Children<Newtonsoft.Json.Linq.JObject>())
            {
                UpdateInfo info = new UpdateInfo();
                info.TagName = (string)release["tag_name"];
                info.Name = (string)release["name"];
                info.IsPrerelease = (bool?)release["prerelease"] == true;
                string htmlUrl = (string)release["html_url"];
                info.NotesUrl = htmlUrl;
                info.Version = ParseVersion(info.TagName) ?? ParseVersion(info.Name);
                Newtonsoft.Json.Linq.JArray assets = release["assets"] as Newtonsoft.Json.Linq.JArray;
                if (assets != null)
                {
                    foreach (Newtonsoft.Json.Linq.JObject asset in assets.Children<Newtonsoft.Json.Linq.JObject>())
                    {
                        string assetName = (string)asset["name"];
                        if (assetName != null && assetName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)
                            && assetName.IndexOf("installer", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            info.InstallerUrl = (string)asset["browser_download_url"];
                            info.InstallerFileName = assetName;
                            info.InstallerSize = (long?)asset["size"] ?? 0;
                            break;
                        }
                    }
                }
                releases.Add(info);
            }
            return releases;
        }

        /// <summary>
        /// Extract a x.y.z.w version from a release tag or name
        /// </summary>
        private static Version ParseVersion(string tagOrName)
        {
            if (string.IsNullOrEmpty(tagOrName))
            {
                return null;
            }
            foreach (string part in System.Text.RegularExpressions.Regex.Split(tagOrName, "[^0-9\\.]+"))
            {
                if (string.IsNullOrEmpty(part))
                {
                    continue;
                }
                string candidate = part.Trim('.');
                Version version;
                if (Version.TryParse(candidate, out version))
                {
                    return version;
                }
            }
            return null;
        }
    }
}
