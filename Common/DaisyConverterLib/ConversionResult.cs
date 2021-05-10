namespace Daisy.DaisyConverter.DaisyConverterLib
{
	/// <summary>
	/// Provides details about convertation.
	/// </summary>
	public class ConversionResult
	{
		private ConversionResult(ConversionResultType resultType)
		{
			ResultType = resultType;
		}

		/// <summary>
		/// Creates success result.
		/// </summary>
		/// <returns></returns>
		public static ConversionResult Success()
		{
			return new ConversionResult(ConversionResultType.Success);
		}

		/// <summary>
		/// Creates cancel result.
		/// </summary>
		/// <returns></returns>
		public static ConversionResult Cancel()
		{
			return new ConversionResult(ConversionResultType.Cancel);
		}

		/// <summary>
		/// Creates invalid result.
		/// </summary>
		/// <param name="error"></param>
		/// <returns></returns>
		public static ConversionResult ValidationError(string error)
		{
			var result = new ConversionResult(ConversionResultType.ValidationError);
			result.ValidationErrorMessage = error;
			return result;
		}

		/// <summary>
		/// Creates UnknownError result.
		/// </summary>
		/// <param name="errorMessage"></param>
		/// <returns></returns>
		public static ConversionResult UnknownError(string errorMessage)
		{
			var result = new ConversionResult(ConversionResultType.UnknownError);
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
		public ConversionResultType ResultType { get; set; }
	}
}