using log4net.Appender;
using log4net.Core;
using System;

namespace Daisy.SaveAsDAISY.Conversion.Telemetry
{
    /// <summary>
    /// log4net appender forwarding the addin log entries to the OpenPanel
    /// telemetry instance, so that logs and exceptions can be monitored there.
    /// </summary>
    public class OpenPanelAppender : AppenderSkeleton
    {
        protected override void Append(LoggingEvent loggingEvent)
        {
            string message = loggingEvent.RenderedMessage;
            Exception exception = loggingEvent.ExceptionObject;
            OpenPanelTelemetry.TrackLog(
                loggingEvent.Level?.Name ?? "INFO",
                message,
                exception
            );
        }
    }
}
