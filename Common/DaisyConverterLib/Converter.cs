using Daisy.SaveAsDAISY.Conversion.Events;
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

namespace Daisy.SaveAsDAISY.Conversion {

    /// <summary>
    /// Base class to convert one or more preprocessed document to DTBook XML
    /// and then possibly to other format using the Daisy pipeline
    /// </summary>
    public class Converter {

		/// <summary>
		/// 
		/// </summary>
		protected WordToDTBookXMLTransform documentConverter;
		protected ConversionParameters conversion;

		protected ChainResourceManager resourceManager;

		protected string validationErrorMsg = "";

		protected int flag;

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

		/// <summary>
		/// Events handler class
		/// </summary>
		protected IConversionEventsHandler eventsHandler;

		public Converter(IConversionEventsHandler eventsHandler, WordToDTBookXMLTransform documentConverter, ConversionParameters conversionParameters) {
			this.eventsHandler = eventsHandler;
			this.conversion = conversionParameters;
			this.documentConverter = documentConverter;
			this.documentConverter.RemoveMessageListeners();
			this.documentConverter.AddProgressMessageListener(new WordToDTBookXMLTransform.XSLTMessagesListener(this.eventsHandler.onProgressMessageReceived));
			this.documentConverter.AddFeedbackMessageListener(new WordToDTBookXMLTransform.XSLTMessagesListener(this.eventsHandler.onFeedbackMessageReceived));
			this.documentConverter.AddFeedbackValidationListener(new WordToDTBookXMLTransform.XSLTMessagesListener(this.eventsHandler.onFeedbackValidationMessageReceived));
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



        /// <summary>
        /// Convert a document to XML and apply post processing script on the result
        /// (Note that post-processing can be a conversion to another format)
        /// </summary>
        /// <param name="document"></param>
        /// <param name="conversion"></param>
        /// <param name="applyPostProcessing"></param>
        public ConversionResult convert(DocumentParameters document, bool applyPostProcessing = true) {

			this.eventsHandler.onDocumentConversionStart(document, conversion);
			xmlConversionCancel = new CancellationTokenSource();
			isCanceled = false;

			if (document.SubDocuments.Count > 0 && conversion.ParseSubDocuments.ToLower() == "yes") {
                List<DocumentParameters> flattenList = new List<DocumentParameters> {
                    document
                };
                foreach (DocumentParameters subDocument in document.SubDocuments) {
					flattenList.Add(subDocument);
				}
				this.convert(flattenList, false);
			} else {
				try {
					conversionTask = Task<XmlDocument>.Factory.StartNew(() => {
						return documentConverter.ConvertDocument(document, conversion);
					});
					conversionTask.Wait(xmlConversionCancel.Token);
				} catch (OperationCanceledException) { 
				} catch (Exception e) {
					this.eventsHandler.OnUnknownError(e.Message);
					return ConversionResult.UnknownError(e.Message);
				}
				if (conversionTask.IsFaulted) {
					this.eventsHandler.OnUnknownError(conversionTask.Exception.Message);
					return ConversionResult.UnknownError(conversionTask.Exception.Message);
				}
			}
			if (isCanceled) { // Conversion is aborted
				this.eventsHandler.onConversionCanceled();
				return ConversionResult.Cancel();
			}

			this.eventsHandler.onDocumentConversionSuccess(document, conversion);
			if (applyPostProcessing && conversion.PostProcessSettings != null) { // launch the pipeline post processing
				this.eventsHandler.onPostProcessingStart(conversion);
				try {
					conversion.PostProcessSettings.ExecuteScript(conversion.OutputPath);
				} catch (Exception e) {
					this.eventsHandler.OnUnknownError(e.Message);
					return ConversionResult.UnknownError(e.Message);
				}
				this.eventsHandler.onPostProcessingSuccess(conversion);
			}
			return ConversionResult.Success();
		}

		/// <summary>
		/// Convert a list of document, merge them in a single XML file and apply post processing on the merged document
		/// </summary>
		/// <param name="documentLists">list of one or more document to convert</param>
		/// <param name="conversion">global conversion settings</param>
		/// <param name="applyPostProcessing">if true, post processing will be applied on the merge result</param>
		public ConversionResult convert(List<DocumentParameters> documentLists, bool applyPostProcessing = true) {
			this.eventsHandler.onDocumentListConversionStart(documentLists, conversion);
			isCanceled = false;
			string errors = "";

			string outputDirectory = conversion.OutputPath.EndsWith(".xml") ?
					Directory.GetParent(conversion.OutputPath).FullName :
					conversion.OutputPath;
			string outputFilename = conversion.OutputPath.EndsWith(".xml") ?
					Path.GetFileName(conversion.OutputPath) :
					Path.GetFileNameWithoutExtension(documentLists[0].InputPath) + ".xml";

			// Rebuild output path
			conversion.OutputPath = Path.Combine(outputDirectory, outputFilename);

			try {
				XmlDocument mergeResult = new XmlDocument();
				string executionErrors = "";
				if (documentLists.Count == 1) {
					foreach (DocumentParameters document in documentLists) {
						if (isCanceled) return ConversionResult.Cancel();
						//onProgressMessageReceived(this, new DaisyEventArgs("Converting " + document.InputPath));
						//document.OutputPath = outputDirectory + "\\" + Path.GetFileNameWithoutExtension(document.InputPath) + ".xml";
						this.convert(document, false);
					}
				} else {
					
					foreach (DocumentParameters document in documentLists) {
						
						document.OutputPath = outputDirectory + "\\" + Path.GetFileNameWithoutExtension(document.InputPath) + ".xml";
                        try {
							
							conversionTask = Task<XmlDocument>.Factory.StartNew(() => {
								return documentConverter.ConvertDocument(document, conversion, mergeResult);
							});

							conversionTask.Wait(xmlConversionCancel.Token);
							if (conversionTask.IsCanceled) {
								this.eventsHandler.onConversionCanceled();
								return ConversionResult.Cancel();
							} else {
								mergeResult = conversionTask.Result;
							}
						} catch (Exception e) {
							executionErrors += document.InputPath + ": " + e.Message + "\r\n";
						}

						if (conversionTask.IsFaulted) {
							this.eventsHandler.OnUnknownError(conversionTask.Exception.Message);
							return ConversionResult.UnknownError(conversionTask.Exception.Message);
						}

						if (isCanceled) {
							this.eventsHandler.onConversionCanceled();
							return ConversionResult.Cancel();
						}

					}
					documentConverter.finalizeAndSaveMergedDocument(mergeResult, conversion);
				}

			} catch (Exception e) {
				// Propagate unhandled exception
				throw new Exception(resourceManager.GetString("TranslationFailed") + "\n"
					+ resourceManager.GetString("WellDaisyFormat") + "\n" + " \""
					+ errors + "\n" + "Crictical issue:" + "\n" + e.Message + "\n");

			}
			if (documentConverter.ValidationErrors.Count > 0) {
				this.eventsHandler.OnValidationErrors(documentConverter.ValidationErrors, conversion.OutputPath );
				return ConversionResult.ValidationError(
					string.Join(
						"\r\n",
						documentConverter.ValidationErrors.Select(
							error => error.ToString()
						).ToArray()
					)
				);
			} else if (isCanceled) {
				return ConversionResult.Cancel();
			} else {
				this.eventsHandler.onDocumentListConversionSuccess(documentLists, conversion);
				
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
					this.eventsHandler.OnLostElements(conversion.OutputPath, unconvertedElements);
					applyPostProcessing = this.eventsHandler.IsContinueDTBookGenerationOnLostElements();
				}
				if (applyPostProcessing && conversion.PostProcessSettings != null) {
					// If post processing is requested (and can be applied even if lost elements are found)
					// Launch the post processing pipeline sript 
					// (cleaning or converting DTBook to another format Like a DAISY book)
					this.eventsHandler.onPostProcessingStart(conversion);
					try {
						conversion.PostProcessSettings.ExecuteScript(conversion.OutputPath);
					} catch (Exception e) {
						this.eventsHandler.OnUnknownError(e.Message);
						return ConversionResult.UnknownError(e.Message);
					}
					this.eventsHandler.onPostProcessingSuccess(conversion);
				}
			}
			return ConversionResult.Success();

		}


		/// <summary>
		/// 
		/// </summary>
		protected void requestConversionCancel() {
			if (xmlConversionCancel != null) {
				xmlConversionCancel.Cancel();
			}
			isCanceled = false;
		}


	}
}
