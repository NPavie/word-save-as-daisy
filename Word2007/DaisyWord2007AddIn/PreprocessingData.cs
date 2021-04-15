using System.Collections;
using Daisy.DaisyConverter.DaisyConverterLib;

namespace DaisyWord2007AddIn
{
	public class PreprocessingData
	{
		public PreprocessingData()
		{
			ObjectShapes = new ArrayList();
			ImageId = new ArrayList();
			InlineShapes = new ArrayList();
			InlineId = new ArrayList();
			
			MathMLEquations = new Hashtable();
			MultipleOwnMathMl = new Hashtable();
		}

		public static PreprocessingData Failed(string error)
		{
			return new PreprocessingData() { IsSuccess = false, IsCanceled = false, LastMessage = error };
		}

		public static PreprocessingData Canceled(string message) {
			return new PreprocessingData() { IsSuccess = false, IsCanceled = true, LastMessage = message };
		}

		public bool IsSuccess { get; set; }

		public bool IsCanceled { get; set; }

		public string LastMessage { get; set; }
		public string OriginalFilePath { get; set; }
		public string TempFilePath { get; set; }
		public Initialize InitializeWindow { get; set; }
		public string MasterSubFlag { get; set; }
		public ArrayList ObjectShapes { get; set; }
		public ArrayList ImageId { get; set; }
		public ArrayList InlineShapes { get; set; }
		public ArrayList InlineId { get; set; }
		public Hashtable MathMLEquations { get; set; }
		public Hashtable MultipleOwnMathMl { get; set; }
	}
}