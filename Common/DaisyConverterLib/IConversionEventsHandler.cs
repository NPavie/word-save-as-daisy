using System;
using System.Windows.Forms;

namespace Daisy.SaveAsDAISY.DaisyConverterLib
{
	public interface IConversionEventsHandler
	{
		void OnStop(string message);
		bool AskForTranslatingSubdocuments();
		void OnError(string errorMessage);
		void OnStop(string message, string title);

	}

	public class ConsoleEventsHandler : IConversionEventsHandler
	{
		public void OnStop(string message)
		{
			Console.WriteLine("ERROR : "  + message);
		}

		public bool AskForTranslatingSubdocuments()
		{
			return false;
		}

		public void OnError(string errorMessage)
		{
			Console.WriteLine("ERROR : " + errorMessage);
		}

		public void OnStop(string message, string title)
		{
			OnStop(message);
		}
	}

}