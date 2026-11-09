using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Daisy.SaveAsDAISY.Conversion.Telemetry
{
    /// <summary>
    /// Reports usage telemetry of the addin to a self-hosted OpenPanel instance
    /// (https://openpanel.dev). Reports are fire-and-forget: failures are silently
    /// ignored to never impact the addin behavior.
    /// Configuration is read from a "telemetry.xml" file in the addin APPDATA
    /// directory (see telemetry/openpanel/README.md at the repository root).
    /// </summary>
    public static class OpenPanelTelemetry
    {
        private const string ConfigFileName = "telemetry.xml";
        private const string DefaultEventSource = "word-save-as-daisy";

        private static readonly object configLock = new object();
        private static bool configLoaded;
        private static bool enabled;
        private static string apiUrl;
        private static string clientId;
        private static string clientSecret;
        private static string profileId;

        private static HttpClient client;

        #region configuration

        private static void LoadConfig()
        {
            lock (configLock)
            {
                if (configLoaded)
                {
                    return;
                }
                configLoaded = true;

                try
                {
                    string configFile = Path.Combine(
                        ConverterHelper.AppDataSaveAsDAISYDirectory,
                        ConfigFileName
                    );
                    if (!File.Exists(configFile))
                    {
                        return;
                    }

                    XmlDocument doc = new XmlDocument();
                    doc.Load(configFile);
                    XmlNode root = doc.SelectSingleNode("telemetry") ?? doc.DocumentElement;

                    enabled = GetBool(root, "enabled");
                    apiUrl = GetString(root, "apiUrl");
                    clientId = GetString(root, "clientId");
                    clientSecret = GetString(root, "clientSecret");

                    if (enabled && !string.IsNullOrEmpty(apiUrl) && !string.IsNullOrEmpty(clientId))
                    {
                        profileId = GetOrCreateProfileId(root, configFile, doc);
                        InitializeClient();
                    }
                    else
                    {
                        enabled = false;
                    }
                }
                catch
                {
                    // telemetry must never break the addin
                    enabled = false;
                }
            }
        }

        private static void InitializeClient()
        {
            client = new HttpClient();
            client.Timeout = TimeSpan.FromSeconds(10);
            client.DefaultRequestHeaders.Add("openpanel-client-id", clientId);
            if (!string.IsNullOrEmpty(clientSecret))
            {
                client.DefaultRequestHeaders.Add("openpanel-client-secret", clientSecret);
            }
            client.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/json")
            );
        }

        private static string GetString(XmlNode root, string name)
        {
            XmlNode node = root?.SelectSingleNode(name);
            return node?.InnerText?.Trim();
        }

        private static bool GetBool(XmlNode root, string name)
        {
            bool value;
            return bool.TryParse(GetString(root, name), out value) && value;
        }

        /// <summary>
        /// Anonymous but stable identifier of the installation, used by OpenPanel
        /// to group events per user. Stored in the telemetry config file.
        /// </summary>
        private static string GetOrCreateProfileId(XmlNode root, string configFile, XmlDocument doc)
        {
            string id = GetString(root, "profileId");
            if (!string.IsNullOrEmpty(id))
            {
                return id;
            }

            using (SHA256 sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(
                    Encoding.UTF8.GetBytes(
                        Environment.UserDomainName + "\\" + Environment.UserName
                    )
                );
                id = "user-" + BitConverter.ToString(hash)
                    .Replace("-", "")
                    .ToLowerInvariant()
                    .Substring(0, 16);
            }

            XmlElement node = doc.CreateElement("profileId");
            node.InnerText = id;
            root.AppendChild(node);
            doc.Save(configFile);
            return id;
        }

        #endregion

        #region public tracking API

        /// <summary>
        /// Track the usage of an addin UI action (ribbon button, dialog...).
        /// </summary>
        public static void TrackAction(string action)
        {
            Track("addin_action", new Dictionary<string, object>
            {
                { "action", action },
                { "source", DefaultEventSource }
            });
        }

        /// <summary>
        /// Track the successful conversion of a document.
        /// </summary>
        public static void TrackConversion(ConversionParameters conversion)
        {
            Dictionary<string, object> properties = new Dictionary<string, object>
            {
                { "source", DefaultEventSource }
            };
            if (conversion?.PipelineScript != null)
            {
                properties["script"] = conversion.PipelineScript.Name;
                properties["scriptNiceName"] = conversion.PipelineScript.NiceName;
            }
            Track("document_converted", properties);
        }

        /// <summary>
        /// Track a log entry forwarded from the addin logger.
        /// </summary>
        public static void TrackLog(string level, string message, Exception exception)
        {
            Dictionary<string, object> properties = new Dictionary<string, object>
            {
                { "source", DefaultEventSource },
                { "level", level },
                { "message", message }
            };
            if (exception != null)
            {
                if (exception is AggregateException aggregate)
                {
                    exception = aggregate.InnerException ?? exception;
                }
                properties["exceptionType"] = exception.GetType().FullName;
                properties["exceptionMessage"] = exception.Message;
                properties["stackTrace"] = exception.StackTrace;
            }
            Track("addin_log", properties);
        }

        /// <summary>
        /// Track an unexpected exception reported to the user.
        /// </summary>
        public static void TrackException(Exception exception, string context)
        {
            if (exception is AggregateException aggregate)
            {
                exception = aggregate.InnerException ?? exception;
            }
            Track("addin_exception", new Dictionary<string, object>
            {
                { "source", DefaultEventSource },
                { "context", context },
                { "exceptionType", exception.GetType().FullName },
                { "exceptionMessage", exception.Message },
                { "stackTrace", exception.StackTrace }
            });
        }

        /// <summary>
        /// Send a track event to the OpenPanel /track endpoint.
        /// </summary>
        public static void Track(string eventName, Dictionary<string, object> properties)
        {
            try
            {
                LoadConfig();
                if (!enabled)
                {
                    return;
                }

                Dictionary<string, object> payload = new Dictionary<string, object>
                {
                    { "type", "track" },
                    {
                        "payload",
                        new Dictionary<string, object>
                        {
                            { "name", eventName },
                            { "profileId", profileId },
                            { "properties", properties }
                        }
                    }
                };

                Task.Run(async () =>
                {
                    try
                    {
                        StringContent content = new StringContent(
                            JsonConvert.SerializeObject(payload),
                            Encoding.UTF8,
                            "application/json"
                        );
                        using (HttpResponseMessage response = await client
                            .PostAsync(apiUrl.TrimEnd('/') + "/track", content)
                            .ConfigureAwait(false))
                        {
                            // response status is intentionally ignored
                        }
                    }
                    catch
                    {
                        // network failures are silently ignored
                    }
                });
            }
            catch
            {
                // telemetry must never break the addin
            }
        }

        #endregion
    }
}
