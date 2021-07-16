namespace Daisy.SaveAsDAISY.DaisyConverterLib
{
	/// <summary>
	/// Provides details about a conversion.
	/// </summary>
	public class ConversionResult
	{
		public enum Type {
			Success,
			ValidationError,
			Cancel,
			UnknownError
		}


		private ConversionResult(Type resultType)
		{
			ResultType = resultType;
		}

		/// <summary>
		/// Creates success result.
		/// </summary>
		/// <returns></returns>
		public static ConversionResult Success()
		{
			return new ConversionResult(Type.Success);
		}

		/// <summary>
		/// Creates cancel result.
		/// </summary>
		/// <returns></returns>
		public static ConversionResult Cancel()
		{
			return new ConversionResult(Type.Cancel);
		}

		/// <summary>
		/// Creates invalid result.
		/// </summary>
		/// <param name="error"></param>
		/// <returns></returns>
		public static ConversionResult ValidationError(string error)
		{
			var result = new ConversionResult(Type.ValidationError);
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
			var result = new ConversionResult(Type.UnknownError);
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
		public Type ResultType { get; set; }
	}
}