using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Daisy.SaveAsDAISY.DaisyConverterLib {
    public class ConverterHelper {

		/// <summary>
		/// Gets path to pipeline root directory.
		/// </summary>
		public static string PipelinePath {
			get { return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\pipeline-lite-ms"; }
		}

		/// <summary>
		/// Indicates if pipeline exists.
		/// </summary>
		/// <returns></returns>
		public static bool PipelineIsInstalled() {
			return Directory.Exists(PipelinePath);
		}

		/// <summary>
		/// Gets path to the addin directory in AppData.
		/// </summary>
		public static string AppDataSaveAsDAISYDirectory {
			get { return Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\SaveAsDAISY"; }
		}

		public static bool documentIsOpen(string documentPath) {
			try {
				Package pack;
				pack = Package.Open(documentPath, FileMode.Open, FileAccess.ReadWrite);
				pack.Close();
			} catch {
				return true;
			}
			return false;
		}

	}
}
