using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace Daisy.SaveAsDAISY.DaisyConverterLib {

	/// <summary>
	/// Base class to convert one or more preprocessed document to XML or other format
	/// The class should handle the conversion to XML as a background thread
	/// pipeline conversion
	/// </summary>
	public abstract class BaseConverter {

		/// <summary>
		/// Pipeline 1 script to apply on the DTBook XML 
		/// </summary>
		public ScriptParser PipelinePostprocessingScript { get; set; }

		/// <summary>
		/// 
		/// </summary>
		protected WordToDTBookXMLTransform documentConverter;

		protected ChainResourceManager resourceManager;

		protected string validationErrorMsg = "";

		protected int flag;
        private WordToDTBookXMLTransform converter;

		private Task<XmlDocument> conversionTask;
		private CancellationTokenSource xmlConversionCancel;

		protected bool isCanceled = false;

        /// <summary>
        /// Override default resource manager.
        /// </summary>
        public System.Resources.ResourceManager OverrideResourceManager {
			set { this.resourceManager.Add(value); }
		}


		/// <summary>
		/// Retrieve the label associated to the specified key
		/// </summary>
		/// <param name="key"></param>
		/// <returns></returns>
		public string GetString(string key) {
			return this.resourceManager.GetString(key);
		}

		/// <summary>
		/// Resource manager
		/// </summary>
		public System.Resources.ResourceManager ResManager {
			get { return this.resourceManager; }
		}


		public BaseConverter(
			WordToDTBookXMLTransform converter,
			ScriptParser dtbookProcessingPipelineScript
		) {
			PipelinePostprocessingScript = dtbookProcessingPipelineScript;

			this.documentConverter = converter;
			this.documentConverter.RemoveMessageListeners();
			this.documentConverter.AddProgressMessageListener(new WordToDTBookXMLTransform.XSLTMessagesListener(onProgressMessageReceived));
			this.documentConverter.AddFeedbackMessageListener(new WordToDTBookXMLTransform.XSLTMessagesListener(onFeedbackMessageReceived));
			this.documentConverter.AddFeedbackValidationListener(new WordToDTBookXMLTransform.XSLTMessagesListener(onFeedbackValidationMessageReceived));
			this.documentConverter.DirectTransform = true;

			this.resourceManager = new ChainResourceManager();
			// Add a default resource managers (for common labels)
			this.resourceManager.Add(
				new System.Resources.ResourceManager(
					"DaisyAddinLib.resources.Labels",
					Assembly.GetExecutingAssembly()
				)
			);
		}

        protected BaseConverter(WordToDTBookXMLTransform converter) {
            this.converter = converter;
        }




        /// <summary>
        /// Convert a document to XML and apply post processing script on the result
        /// (Note that post-processing can be a conversion to another format)
        /// </summary>
        /// <param name="document"></param>
        /// <param name="conversion"></param>
        /// <param name="applyPostProcessing"></param>
        public ConversionResult convert(DocumentParameters document, ConversionParameters conversion, bool applyPostProcessing = true) {
			
			this.onDocumentConversionStart(document, conversion);
			
			isCanceled = false;

			if (document.SubDocuments.Count > 0 && conversion.ParseSubDocuments.ToLower() == "yes") {
				List<DocumentParameters> flattenList = new List<DocumentParameters>();
				flattenList.Add(document);
				foreach (DocumentParameters subDocument in document.SubDocuments) {
					flattenList.Add(subDocument);
				}
				this.convert(flattenList, conversion, false);
			} else {
				try {
					conversionTask = Task<XmlDocument>.Factory.StartNew(() => {
						return documentConverter.convertDocument(document, conversion);
					});
					conversionTask.Wait(xmlConversionCancel.Token);
				} catch (OperationCanceledException) { 
				} catch (Exception e) {
					OnUnknownError(e.Message);
					return ConversionResult.UnknownError(e.Message);
				}
			}
			if (isCanceled) { // Conversion is aborted
				return ConversionResult.Cancel();
			}
			this.onDocumentConversionSuccess(document, conversion);
			if (applyPostProcessing && PipelinePostprocessingScript != null) { // launch the pipeline post processing
				this.onPostProcessingStart(conversion);
				try {
					PipelinePostprocessingScript.ExecuteScript(conversion.OutputPath);
				} catch (Exception e) {
					OnUnknownError(e.Message);
					return ConversionResult.UnknownError(e.Message);
				}
				this.onPostProcessingSuccess(conversion);
			}
			
			return ConversionResult.Success();
		}

		/// <summary>
		/// Convert a list of document, merge them in a single XML file and apply post processing on the merged document
		/// </summary>
		/// <param name="documentLists">list of one or more document to convert</param>
		/// <param name="conversion">global conversion settings</param>
		/// <param name="applyPostProcessing">if true, post processing will be applied on the merge result</param>
		public ConversionResult convert(List<DocumentParameters> documentLists, ConversionParameters conversion, bool applyPostProcessing = true) {
			this.onDocumentListConversionStart(documentLists, conversion);
			isCanceled = false;
			string errors = "";
			conversion.InputFile = documentLists[0].InputPath;

			string outputDirectory = conversion.OutputPath.EndsWith(".xml") ?
					Directory.GetParent(conversion.OutputPath).FullName :
					conversion.OutputPath;

			// Rebuild output path
			conversion.OutputPath = Path.Combine(
				outputDirectory,
				conversion.OutputPath.EndsWith(".xml") ?
					Path.GetFileName(conversion.OutputPath) :
					Path.GetFileNameWithoutExtension(conversion.InputFile) + ".xml"
				);

			try {
				XmlDocument mergeResult = new XmlDocument();
				ArrayList mergeDocLanguage = new ArrayList();
				ArrayList lostElements = new ArrayList();

				if (documentLists.Count == 1) {
					foreach (DocumentParameters document in documentLists) {
						if (isCanceled) return ConversionResult.Cancel();
						//onProgressMessageReceived(this, new DaisyEventArgs("Converting " + document.InputPath));
						//document.OutputPath = outputDirectory + "\\" + Path.GetFileNameWithoutExtension(document.InputPath) + ".xml";
						this.convert(document, conversion, false);
					}
				} else {
					foreach (DocumentParameters document in documentLists) {
						
						document.OutputPath = outputDirectory + "\\" + Path.GetFileNameWithoutExtension(document.InputPath) + ".xml";
                        try {
							
							conversionTask = Task<XmlDocument>.Factory.StartNew(() => {
								return documentConverter.convertDocument(document, conversion, mergeResult);
							});

							conversionTask.Wait(xmlConversionCancel.Token);
							mergeResult = conversionTask.Result;
						} catch (Exception) { }
						
						if (isCanceled) return ConversionResult.Cancel();

					}
					documentConverter.finalizeAndSaveMergedDocument(mergeResult, conversion, mergeDocLanguage);
				}

			} catch (Exception e) {
				// Propagate unhandled exception
				throw new Exception(resourceManager.GetString("TranslationFailed") + "\n"
					+ resourceManager.GetString("WellDaisyFormat") + "\n" + " \""
					+ errors + "\n" + "Crictical issue:" + "\n" + e.Message + "\n");

			}
			if (errors.Length > 0) {
				OnValidationError(errors, conversion.InputFile, conversion.OutputPath);
				return ConversionResult.ValidationError(errors);
			} else if (isCanceled) {
				return ConversionResult.Cancel();
			} else {
				onDocumentListConversionSuccess(documentLists, conversion);
				
				if (documentConverter.LostElements.Count > 0) {
					ArrayList unconvertedElements = new ArrayList();
                    foreach (KeyValuePair<string, List<string> > lostElementForFile in documentConverter.LostElements) {
						if(lostElementForFile.Value.Count > 0) {
							string lostElements = lostElementForFile.Key + ":\r\n";
                            foreach (string lostElement in lostElementForFile.Value) {
								lostElements += " - " + lostElement + "\r\n";
							}
							unconvertedElements.Add(lostElements);
                        }
                    }
					OnLostElements(conversion.OutputPath, unconvertedElements);
					applyPostProcessing = IsContinueDTBookGenerationOnLostElements();
				}
				if (applyPostProcessing && PipelinePostprocessingScript != null) {
					// If post processing is requested (and can be applied even if lost elements are found)
					// Launch the post processing pipeline sript 
					// (cleaning or converting DTBook to another format Like a DAISY book)
					this.onPostProcessingStart(conversion);
					try {
						PipelinePostprocessingScript.ExecuteScript(conversion.OutputPath);
					} catch (Exception e) {
						OnUnknownError(e.Message);
						return ConversionResult.UnknownError(e.Message);
					}
					this.onPostProcessingSuccess(conversion);
				}
			}
			return ConversionResult.Success();

		}


        #region virtual methods to override

        #region Stepping functions

		/// <summary>
		/// Method called when the conversion of a list of document starts
		/// </summary>
		/// <param name="documentLists"></param>
		/// <param name="conversion"></param>
        protected virtual void onDocumentListConversionStart(List<DocumentParameters> documentLists, ConversionParameters conversion) {
		}

		/// <summary>
		/// Function called when the conversion of a document starts
		/// </summary>
		/// <param name="document"></param>
		/// <param name="conversion"></param>
		protected virtual void onDocumentConversionStart(DocumentParameters document, ConversionParameters conversion) {
			xmlConversionCancel = new CancellationTokenSource();

		}

		/// <summary>
		/// Method called when post processing starts
		/// </summary>
		/// <param name="conversion"></param>
		protected virtual void onPostProcessingStart(ConversionParameters conversion) {
		}

		/// <summary>
		/// Method called when the conversion of a word documents list to the dtbook xml is successful (before post-processing)
		/// </summary>
		/// <param name="documentLists"></param>
		/// <param name="conversion"></param>
		protected virtual void onDocumentListConversionSuccess(List<DocumentParameters> documentLists, ConversionParameters conversion) {
		}

		/// <summary>
		/// Method called when the conversion of a word document to the dtbook xml is successful (before post-processing)
		/// </summary>
		/// <param name="document"></param>
		/// <param name="conversion"></param>
		protected virtual void onDocumentConversionSuccess(DocumentParameters document, ConversionParameters conversion) {
		}

		/// <summary>
		/// Method called when the post processing pass has successfully finished 
		/// </summary>
		/// <param name="conversion"></param>
		protected virtual void onPostProcessingSuccess(ConversionParameters conversion) {
		}

		protected virtual void onConversionCancel() {
			if(xmlConversionCancel != null) {
				xmlConversionCancel.Cancel();
			}
			isCanceled = false;
		}
        
		#endregion

        /// <summary>
        /// Progress message should indicate progression on the whole conversion process
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected virtual void onProgressMessageReceived(object sender, EventArgs e) {
			
		}

		/// <summary>
		/// Feedback message should be informations like XSLT informations, warning and errors
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		protected virtual void onFeedbackMessageReceived(object sender, EventArgs e) {
			//string message = ((DaisyEventArgs)e).Message;
			//Console.Write(message);
		}

		/// <summary>
		/// Validation feedback message (most probably validation errors)
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		protected virtual void onFeedbackValidationMessageReceived(object sender, EventArgs e) {
			string message = ((DaisyEventArgs)e).Message;
			//Console.Write(message);
		}

		protected virtual void OnSuccessMasterSubValidation(string message) {

		}

		/// <summary>
		/// Called when an unknown error was raised during conversion
		/// </summary>
		/// <param name="error"></param>
		protected virtual void OnUnknownError(string error) {

		}

		protected virtual void OnUnknownError(string title, string details) {
		}

		/// <summary>
		/// Called when a validation error has occured
		/// </summary>
		/// <param name="error"></param>
		/// <param name="inputFile"></param>
		/// <param name="outputFile"></param>
		protected virtual void OnValidationError(string error, string inputFile, string outputFile) {
		}

		/// <summary>
		/// Called after conversion if lost elements have been fond 
		/// (elements found in a document that will not be converted, like TOC of subdocuments)
		/// </summary>
		/// <param name="outputFile"></param>
		/// <param name="elements"></param>
		protected virtual void OnLostElements(string outputFile, ArrayList elements) {
		}

		/// <summary>
		/// Request user if he wants to continue the conversion with lost elements found
		/// </summary>
		/// <returns></returns>
		protected virtual bool IsContinueDTBookGenerationOnLostElements() {
			return true;
		}

		/// <summary>
		/// Called when conversion has finished
		/// </summary>
		protected virtual void OnSuccess() {
		}

		protected virtual void OnMasterSubValidationError(string error) {

		}

		#endregion


	}
}
