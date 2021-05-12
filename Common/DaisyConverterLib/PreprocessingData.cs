using System.Collections;
using System.IO;

namespace Daisy.SaveAsDAISY.DaisyConverterLib {
	public class PreprocessingData
	{
		public PreprocessingData()
		{}

		public PreprocessingData(string wordVersion,  Pipeline pipeline = null, string pipelineScriptKey = "" ) : base() {
			Settings.Version = wordVersion;
			FileInfo postprocessScriptFile;
			if (pipeline == null) {
				Settings.ScriptPath = null;
			} else if (pipelineScriptKey != "") {
				Settings.ScriptPath = pipeline.ScriptsInfo[pipelineScriptKey].FullName;
				Settings.Directory = string.Empty;
			} else if (pipeline.ScriptsInfo.TryGetValue("_postprocess", out postprocessScriptFile)) {
				// Note : adding a default postprocess script for dtbook pipeline special treatment
				// This script is alledgedly not visible to users
				Settings.ScriptPath = postprocessScriptFile.FullName;
				Settings.Directory = string.Empty;
			} else Settings.ScriptPath = null;
		}

		public static PreprocessingData Failed(string error)
		{
			return new PreprocessingData() { IsSuccess = false, IsCanceled = false, LastMessage = error };
		}

		public static PreprocessingData Canceled(string message) {
			return new PreprocessingData() { IsSuccess = false, IsCanceled = true, LastMessage = message };
		}

		// converter settings, to be initialized when needed
		private ConversionParameters converterSettings = null;
		public ConversionParameters Settings { 
			get {
				if (converterSettings == null) {
					converterSettings = new ConversionParameters();
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