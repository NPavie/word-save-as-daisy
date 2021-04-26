namespace Daisy.DaisyConverter.DaisyConverterLib
{
	/// <summary>
	/// Provides details about convertation.
	/// </summary>
	public class ConverterResult
	{
		private ConverterResult(ConverterResultType resultType)
		{
			ResultType = resultType;
		}

		/// <summary>
		/// Creates success result.
		/// </summary>
		/// <returns></returns>
		public static ConverterResult Success()
		{
			return new ConverterResult(ConverterResultType.Success);
		}

		/// <summary>
		/// Creates cancel result.
		/// </summary>
		/// <returns></returns>
		public static ConverterResult Cancel()
		{
			return new ConverterResult(ConverterResultType.Cancel);
		}

		/// <summary>
		/// Creates invalid result.
		/// </summary>
		/// <param name="error"></param>
		/// <returns></returns>
		public static ConverterResult ValidationError(string error)
		{
			var result = new ConverterResult(ConverterResultType.ValidationError);
			result.ValidationErrorMessage = error;
			return result;
		}

		/// <summary>
		/// Creates UnknownError result.
		/// </summary>
		/// <param name="errorMessage"></param>
		/// <returns></returns>
		public static ConverterResult UnknownError(string errorMessage)
		{
			var result = new ConverterResult(ConverterResultType.UnknownError);
			result.UnknownErrorMessage = errorMessage;
			return result;
		}

		/// <summary>
		/// Error message.
		/// </summary>
		public string UnknownErrorMessage { get; set; }

		/// <summary>
		/// Validation error message.
		/// </summary>
		public string ValidationErrorMessage { get; set; }

		/// <summary>
		/// Result type.
		/// </summary>
		public ConverterResultType ResultType { get; set; }
	}
}