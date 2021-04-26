using System.Collections;
using System.IO;
using Daisy.DaisyConverter.DaisyConverterLib;

namespace DaisyWord2007AddIn
{
	public class ProcessingData
	{
		public ProcessingData()
		{}

		public ProcessingData(string wordVersion, string conversionMode = "", Pipeline scripts = null ) : base() {
			Settings.Version = wordVersion;
			FileInfo postprocessScriptFile;
			if (scripts == null) {
				Settings.ScriptPath = null;
			} else if (!AddInHelper.buttonIsSingleWordToXMLConversion(conversionMode)) {
				Settings.ScriptPath = scripts.ScriptsInfo[conversionMode].FullName;
				Settings.Directory = string.Empty;
			} else if (scripts.ScriptsInfo.TryGetValue("_postprocess", out postprocessScriptFile)) {
				// Note : adding a postprocess script for dtbook pipeline special treatment
				Settings.ScriptPath = postprocessScriptFile.FullName;
				Settings.Directory = string.Empty;
			} else Settings.ScriptPath = null;
		}

		public static ProcessingData Failed(string error)
		{
			return new ProcessingData() { IsSuccess = false, IsCanceled = false, LastMessage = error };
		}

		public static ProcessingData Canceled(string message) {
			return new ProcessingData() { IsSuccess = false, IsCanceled = true, LastMessage = message };
		}

		// converter settings, to be initialized when needed
		private ConverterParameters converterSettings = null;
		public ConverterParameters Settings { 
			get {
				if (converterSettings == null) {
					converterSettings = new ConverterParameters();
					converterSettings.ListMathMl = new Hashtable();
					converterSettings.ObjectShapes = new ArrayList();
					converterSettings.ImageIds = new ArrayList();
					converterSettings.InlineShapes = new ArrayList();
					converterSettings.InlineIds = new ArrayList();
				}
				return converterSettings;
			} 
		}

		public bool IsSuccess { get; set; }

		public bool IsCanceled { get; set; }
		public string LastMessage { get; set; }
	}
}