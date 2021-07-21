namespace Daisy.SaveAsDAISY.Conversion
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
            return new ConversionResult(Type.ValidationError) {
				ValidationErrorMessage = error
			};
		}

		/// <summary>
		/// Creates UnknownError result.
		/// </summary>
		/// <param name="errorMessage"></param>
		/// <returns></returns>
		public static ConversionResult UnknownError(string errorMessage)
		{
            return new ConversionResult(Type.UnknownError) {
				UnknownErrorMessage = errorMessage
			};
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