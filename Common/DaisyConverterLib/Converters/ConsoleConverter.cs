using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Reflection;
using System.Resources;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace Daisy.SaveAsDAISY.DaisyConverterLib
{
    /// <summary>
    /// Full conversion in console mode
    /// </summary>
    public class ConsoleConverter : BaseConverter {
		

		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="converter">An implementation of AbstractConverter</param>
		public ConsoleConverter(WordToDTBookXMLTransform converter)
			: base(converter, null)
		{

		}

		public ConsoleConverter(WordToDTBookXMLTransform converter, ScriptParser scriptToExecute) : base(converter, scriptToExecute)
		{
			
		}

        protected override bool IsContinueDTBookGenerationOnLostElements() {
            return base.IsContinueDTBookGenerationOnLostElements();
        }

        protected override void onConversionCancel() {
            base.onConversionCancel();
        }

        protected override void onDocumentConversionStart(DocumentParameters document, ConversionParameters conversion) {
            base.onDocumentConversionStart(document, conversion);
        }

        protected override void onDocumentConversionSuccess(DocumentParameters document, ConversionParameters conversion) {
            base.onDocumentConversionSuccess(document, conversion);
        }

        protected override void onDocumentListConversionStart(List<DocumentParameters> documentLists, ConversionParameters conversion) {
            base.onDocumentListConversionStart(documentLists, conversion);
        }

        protected override void onDocumentListConversionSuccess(List<DocumentParameters> documentLists, ConversionParameters conversion) {
            base.onDocumentListConversionSuccess(documentLists, conversion);
        }

        protected override void onFeedbackMessageReceived(object sender, EventArgs e) {
            base.onFeedbackMessageReceived(sender, e);
        }

        protected override void onFeedbackValidationMessageReceived(object sender, EventArgs e) {
            base.onFeedbackValidationMessageReceived(sender, e);
        }

        protected override void OnLostElements(string outputFile, ArrayList elements) {
            base.OnLostElements(outputFile, elements);
        }

        protected override void OnMasterSubValidationError(string error) {
            base.OnMasterSubValidationError(error);
        }

        protected override void onPostProcessingStart(ConversionParameters conversion) {
            base.onPostProcessingStart(conversion);
        }

        protected override void onPostProcessingSuccess(ConversionParameters conversion) {
            base.onPostProcessingSuccess(conversion);
        }

        protected override void onProgressMessageReceived(object sender, EventArgs e) {
            base.onProgressMessageReceived(sender, e);
        }

        protected override void OnSuccess() {
            base.OnSuccess();
        }

        protected override void OnSuccessMasterSubValidation(string message) {
            base.OnSuccessMasterSubValidation(message);
        }

        protected override void OnUnknownError(string error) {
            base.OnUnknownError(error);
        }

        protected override void OnUnknownError(string title, string details) {
            base.OnUnknownError(title, details);
        }

        protected override void OnValidationError(string error, string inputFile, string outputFile) {
            base.OnValidationError(error, inputFile, outputFile);
        }
    }
}