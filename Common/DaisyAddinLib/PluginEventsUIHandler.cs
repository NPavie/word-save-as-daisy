using Daisy.SaveAsDAISY.DaisyConverterLib;
using System;
using System.Windows.Forms;

namespace Daisy.SaveAsDAISY
{

	public class PluginEventsUIHandler : IConversionEventsHandler
	{
		public void OnStop(string message)
		{
			OnStop(message,"SaveAsDAISY");
		}

		public bool AskForTranslatingSubdocuments()
		{
			DialogResult dialogResult = MessageBox.Show("Do you want to translate the current document along with sub documents?", "SaveAsDAISY", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
			return dialogResult == DialogResult.Yes;
		}

		public void OnError(string errorMessage)
		{
			MessageBox.Show(errorMessage, "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Error);
		}

		public void OnStop(string message, string title)
		{
			MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Stop);
		}
	}
}