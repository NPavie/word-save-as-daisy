using System;
using System.Collections;
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
	/// Full conversion in quiet mode
	/// </summary>
	public class SingleConverter
	{
		public ScriptParser ScriptToExecute { get; set; }
		private WordToDTBookXMLConverter converter;
		ChainResourceManager resourceManager;
		string validationErrorMsg = "";
		private bool continueDTBookGeneration = true;
		int flag;

		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="converter">An implementation of AbstractConverter</param>
		public SingleConverter(WordToDTBookXMLConverter converter)
			: this(converter, null)
		{

		}

		public SingleConverter(WordToDTBookXMLConverter converter, ScriptParser scriptToExecute)
		{
			ScriptToExecute = scriptToExecute;
			this.converter = converter;
			this.resourceManager = new ChainResourceManager();
			// Add a default resource managers (for common labels)
			this.resourceManager.Add(new System.Resources.ResourceManager("DaisyAddinLib.resources.Labels",
				Assembly.GetExecutingAssembly()));
		}

		/// <summary>
		/// Override default resource manager.
		/// </summary>
		public System.Resources.ResourceManager OverrideResourceManager
		{
			set { this.resourceManager.Add(value); }
		}


		/// <summary>
		/// Retrieve the label associated to the specified key
		/// </summary>
		/// <param name="key"></param>
		/// <returns></returns>
		public string GetString(string key)
		{
			return this.resourceManager.GetString(key);
		}

		/// <summary>
		/// Resource manager
		/// </summary>
		public System.Resources.ResourceManager ResManager
		{
			get { return this.resourceManager; }
		}

		/// <summary>
		/// Function to translate all Sub documents 
		/// </summary>
		/// <param name="tempInputFile">Duplicate for the Input File</param>
		/// <param name="temp_InputFile">Duplicate for the Input File</param>
		/// <param name="inputFile">Input file to be Translated</param>
		/// <param name="outputfilepath">Path of the Output file</param>
		/// <param name="HTable">Document properties</param>
		public void OoxToDaisyOwn(String tempInputFile, String inputFile, String outputfilepath, Hashtable HTable, string control, Hashtable listMathMl, string output_Pipeline)
		{
			try
			{
				SubdocumentsList subdocuments = SubdocumentsManager.FindSubdocuments(
					tempInputFile,
					inputFile);

				if (subdocuments.Errors.Count > 0) {
					StringBuilder errorMessage = new StringBuilder();
					errorMessage.Append("Errors were encoutered while retrieving sub documents:");
					foreach (string error in subdocuments.Errors) {
						errorMessage.Append("\r\n- " + error);
					}
					OnUnknownError(errorMessage.ToString());
					return;
				}
				
				ArrayList subList = new ArrayList();
				subList.Add(tempInputFile + "|Master");
				foreach (string subdoc in subdocuments.GetSubdocumentsNameWithRelationship()) {
					subList.Add(subdoc);
				}

				int subCount = subdocuments.SubdocumentsCount + 1;
				//Checking whether any original or Subdocumets is already Open or not
				foreach (string docPathAndType in subList) {
					string[] splitted = docPathAndType.Split('|');
					if (ConverterHelper.documentIsOpen(splitted[0])) {
						OnUnknownError(resourceManager.GetString("OpenSubOwn"));
						return;
					}
				}

				//Checking whether Sub documents are Simple documents or a Master document
				string resultSub = SubdocumentsManager.CheckingSubDocs(subdocuments.GetSubdocumentsNameWithRelationship());
				if (resultSub != "simple") {
					OnUnknownError(resourceManager.GetString("AddSimpleMasterSub"));
					return;
				}

				// For each document

				OoxToDaisySub(outputfilepath, subList, HTable, tempInputFile, control, listMathMl, output_Pipeline,
							  subdocuments.GetNotTraslatedSubdocumentsNames());

			}
			catch (Exception e)
			{
				OnUnknownError(resourceManager.GetString("TranslationFailed") + "\n" + this.resourceManager.GetString("WellDaisyFormat") + "\n" + " \"" + Path.GetFileName(tempInputFile) + "\"\n" + validationErrorMsg + "\n" + "Problem is:" + "\n" + e.Message + "\n");
			}
		}

		/// <summary>
		/// Function to translate all Bunch of documents  selected by the user
		/// </summary>
		/// <param name="outputfilepath">Path of the output File</param>
		/// <param name="subList">List of All documents</param>
		/// <param name="category">whether master/sub doc or Bunch of Docs</param>
		/// <param name="table">Document Properties</param>
		public bool OoxToDaisySub(
			string outputfilepath,
			ArrayList subList,
			string category,
			Hashtable table,
			string control,
			Hashtable MultipleMathMl,
			string output_Pipeline
		) {
			flag = 0;
			validationErrorMsg = "";
			using (Progress progress = new Progress(this.converter, this.resourceManager, outputfilepath, subList, category, table, control, MultipleMathMl, output_Pipeline))
			{
				DialogResult progressDialogResult = progress.ShowDialog();
				if (progressDialogResult == DialogResult.OK)
				{
					if (!string.IsNullOrEmpty(progress.ValidationError))
					{
						validationErrorMsg = progress.ValidationError;
						OnMasterSubValidationError(progress.ValidationError);
					}
					else if (progress.HasLostElements)
					{
						OnLostElements(string.Empty, outputfilepath, progress.LostElements);
						flag = 1;
						if (!(AddInHelper.IsSingleDaisyFromMultipleButton(control) && ScriptToExecute == null))
							continueDTBookGeneration = IsContinueDTBookGenerationOnLostElements();
					}
					else
					{
						if (AddInHelper.IsSingleDaisyFromMultipleButton(control) && ScriptToExecute == null)
						{
							OnSuccess();
						}
						flag = 1;
					}
				}
				else if (progressDialogResult == DialogResult.Cancel)
				{
					DeleteDTBookFilesIfExists(outputfilepath);
				}
				else if (!string.IsNullOrEmpty(progress.ValidationError))
				{
					OnMasterSubValidationError(progress.ValidationError);
				}
			}

			DeleteTemporaryImages();

			return (flag == 1) && continueDTBookGeneration;
		}
		
		#region conversion froom "Progress"

		private void onProgressMessageReceived(object sender, EventArgs e) {
        }
		private void onFeedbackMessageReceived(object sender, EventArgs e) {
			string message = ((DaisyEventArgs)e).Message;
			Console.Write(message);
		}


		XmlDocument mergeXmlDoc;
		ArrayList mergeDocLanguage;
		private Exception exception;
		Hashtable table;
		private int size;
		String cutData;
		String tempData = "";
		int subDocFootNum;
		private static bool isValid;
		private bool cancel, converting, computeSize;
		private static string error;
		private ResourceManager manager;
		private static string error_MasterSub = "";
		ArrayList MathList8879, MathList9573, MathListmathml;
		String tempInputFile, outputfilepath, category = "", error_Exception = "", output_Pipeline;
		ArrayList lostElements = new ArrayList();
		String docName = "";
		string control = "";
		String errorText = "";
		string path_For_Pipeline = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\pipeline-lite-ms";
		private Hashtable listMathMl;
		private Hashtable multipleMathMl;

		

		/* Function which Translates the current document along with its sub documents*/
		public void DoTranslation(
			WordToDTBookXMLConverter converter,
			ResourceManager manager,
			String outputfile,
			ArrayList subList, Hashtable table, String inputFile, string control, Hashtable listMathMl, string output_Pipeline) {
			try {
				subDocFootNum = 1;
				error_MasterSub = "";
				
				XmlDocument mergeXmlDoc = new XmlDocument();
				ArrayList mergeDocLanguage = new ArrayList();
				this.computeSize = true;
				converter.RemoveMessageListeners();
				converter.AddProgressMessageListenerMaster(new WordToDTBookXMLConverter.XSLTMessagesListener(onProgressMessageReceived));
				converter.AddFeedbackMessageListener(new WordToDTBookXMLConverter.XSLTMessagesListener(onFeedbackMessageReceived));

				for (int i = 0; i < subList.Count; i++) {
					string[] splt = subList[i].ToString().Split('|');
					docName = splt[0];

					converter.Transform(splt[0], null, table, null, true, "");
				}

				for (int i = 0; i < subList.Count; i++) {
					string[] splt = subList[i].ToString().Split('|');
					String outputFile = outputfilepath + "\\" + Path.GetFileNameWithoutExtension(splt[0]) + ".xml";
					String ridOutputFile = splt[1];
					docName = splt[0];
					this.converter.Transform(splt[0], outputFile, table, (Hashtable)listMathMl["Doc" + i], true, output_Pipeline);
					if (i == 0) {
						ReplaceData(outputFile);
						mergeXmlDoc.Load(outputFile);

						if (File.Exists(outputFile)) {
							File.Delete(outputFile);
						}
					} else {
						ReplaceData(outputFile);
						MergeXml(outputFile, mergeXmlDoc, ridOutputFile, splt[0]);

						if (File.Exists(outputFile)) {
							File.Delete(outputFile);
						}
					}
				}
				SetPageNum(mergeXmlDoc);
				SetImage(mergeXmlDoc);
				SetLanguage(mergeXmlDoc);
				RemoveSubDoc(mergeXmlDoc);
				mergeXmlDoc.Save(outputfilepath + "\\" + Path.GetFileNameWithoutExtension(tempInputFile) + ".xml");
				ReplaceData(outputfilepath + "\\" + Path.GetFileNameWithoutExtension(tempInputFile) + ".xml", true);
				CopyDTDToDestinationfolder(outputfilepath);
				CopyMATHToDestinationfolder(outputfilepath);
				XmlValidation(outputfilepath + "\\" + Path.GetFileNameWithoutExtension(tempInputFile) + ".xml");
				ReplaceData(outputfilepath + "\\" + Path.GetFileNameWithoutExtension(tempInputFile) + ".xml", false);
				if (File.Exists(outputfilepath + "\\dtbook-2005-3.dtd")) {
					File.Delete(outputfilepath + "\\dtbook-2005-3.dtd");
				}
				DeleteMath(outputfilepath, true);
				WorkComplete(null);
			} catch (Exception e) {
				WorkComplete(e);
				error_Exception = manager.GetString("TranslationFailed") + "\n" + manager.GetString("WellDaisyFormat") + "\n" + " \"" + Path.GetFileName(tempInputFile) + "\"\n" + error_MasterSub + "\n" + "Problem is:" + "\n" + e.Message + "\n";

			}
		}


			/// <summary>
			/// Function to translate all sub documents in the Master Document
			/// </summary>
			/// <param name="outputfilepath">Path of the output File</param>
			/// <param name="subList">List of All documents</param>
			/// <param name="HTable">Document Properties</param>
			/// <param name="tempInputFile">Duplicate for the Input File</param>
		public void OoxToDaisySub(
			String outputfilepath,
			ArrayList subList,
			Hashtable HTable,
			String tempInputFile,
			string control,
			Hashtable listMathMl,
			string output_Pipeline,
			ArrayList notTranslatedDoc
		) {
			flag = 0;
			validationErrorMsg = "";
			using (Progress progress = new Progress(this.converter, this.resourceManager, outputfilepath, subList, HTable, tempInputFile, control, listMathMl, output_Pipeline)) {
				DialogResult dr = progress.ShowDialog();
				if (dr == DialogResult.OK) {
					validationErrorMsg = progress.ValidationError;
					String messageDocsSkip = DocumentSkipped(notTranslatedDoc);
					if (!string.IsNullOrEmpty(validationErrorMsg)) {
						validationErrorMsg = validationErrorMsg + messageDocsSkip;
						OnMasterSubValidationError(validationErrorMsg);
					} else if (progress.HasLostElements) {
						OnLostElements(string.Empty, outputfilepath + "\\1.xml", progress.LostElements);

						if (ConverterHelper.PipelineIsInstalled() &&
								AddInHelper.buttonIsSingleWordToXMLConversion(control) &&
								ScriptToExecute != null &&
								IsContinueDTBookGenerationOnLostElements()) {
							try {
								ExecuteScript(outputfilepath + "\\" + Path.GetFileNameWithoutExtension(tempInputFile) + ".xml");
							} catch (Exception e) {
								//AddinLogger.Error(e);
								OnUnknownError(e.Message);
							}
						}
					} else {
						if (!string.IsNullOrEmpty(messageDocsSkip)) {
							OnSuccessMasterSubValidation(ResManager.GetString("SucessLabel") + messageDocsSkip);
						} else {
							if (AddInHelper.IsSingleDaisyTranslate(control) && ScriptToExecute == null) {
								OnSuccess();
							} else ExecuteScript(outputfilepath + "\\" + Path.GetFileNameWithoutExtension(tempInputFile) + ".xml");
						}
					}
				} else if (dr == DialogResult.Cancel) {
					DeleteDTBookFilesIfExists(outputfilepath);
				} else {
					validationErrorMsg = progress.ValidationError;
					if (!string.IsNullOrEmpty(validationErrorMsg)) {
						OnMasterSubValidationError(validationErrorMsg);
					}
				}
			}
			DeleteTemporaryImages();
		} 

		public ConversionResult convertToDaisy(
			string inputFile,
			string outputFile,
			Hashtable listMathMl,
			Hashtable table, string control, string output_Pipeline)
		{
			
			try {
				using (ConversionProgressDialog form = new ConversionProgressDialog(this.converter, inputFile, outputFile, this.resourceManager, true, listMathMl, table, control, output_Pipeline)) {
					if (DialogResult.OK != form.ShowDialog())
						return ConversionResult.Cancel();

					if (!String.IsNullOrEmpty(form.ValidationError)) {
						OnValidationError(form.ValidationError, inputFile, outputFile);
						return ConversionResult.ValidationError(form.ValidationError);
					}

					if (form.HasLostElements) {
						OnLostElements(inputFile, outputFile, form.LostElements);

						if (!(AddInHelper.IsSingleDaisyTranslate(control) && this.ScriptToExecute == null) && IsContinueDTBookGenerationOnLostElements()) {
							ExecuteScript(outputFile);
						}
					} else if (AddInHelper.IsSingleDaisyTranslate(control) && this.ScriptToExecute == null) {
						OnSuccess();
					} else {
						ExecuteScript(outputFile);
					}
				}
			} catch (IOException e) {
				// this is meant to catch "file already accessed by another process", though there's no .NET fine-grain exception for this.
				AddinLogger.Error(e);
				OnUnknownError("UnableToCreateOutputLabel", e.Message);
				return ConversionResult.UnknownError(e.Message + Environment.NewLine + e.StackTrace);
			} catch (Exception e) {
				AddinLogger.Error(e);
				OnUnknownError("DaisyUnexpectedError", e.GetType() + ": " + e.Message + " (" + e.StackTrace + ")");

				if (File.Exists(outputFile)) {
					File.Delete(outputFile);
				}

				return ConversionResult.UnknownError(e.Message + Environment.NewLine + e.StackTrace);
			} finally {
				DeleteTemporaryImages();
			}

			return ConversionResult.Success();
		}

		public ConversionResult convertToDaisy(ConversionParameters parameters) {
			if(parameters.ParseSubDocuments == "Yes") {

            } else { // document has no 

            }
			try {
				using (ConversionProgressDialog form = new ConversionProgressDialog(this.converter, inputFile, outputFile, this.resourceManager, true, listMathMl, table, control, output_Pipeline)) {
					if (DialogResult.OK != form.ShowDialog())
						return ConversionResult.Cancel();

					if (!String.IsNullOrEmpty(form.ValidationError)) {
						OnValidationError(form.ValidationError, inputFile, outputFile);
						return ConversionResult.ValidationError(form.ValidationError);
					}

					if (form.HasLostElements) {
						OnLostElements(inputFile, outputFile, form.LostElements);

						if (!(AddInHelper.IsSingleDaisyTranslate(control) && this.ScriptToExecute == null) && IsContinueDTBookGenerationOnLostElements()) {
							ExecuteScript(outputFile);
						}
					} else if (AddInHelper.IsSingleDaisyTranslate(control) && this.ScriptToExecute == null) {
						OnSuccess();
					} else {
						ExecuteScript(outputFile);
					}
				}
			} catch (IOException e) {
				// this is meant to catch "file already accessed by another process", though there's no .NET fine-grain exception for this.
				AddinLogger.Error(e);
				OnUnknownError("UnableToCreateOutputLabel", e.Message);
				return ConversionResult.UnknownError(e.Message + Environment.NewLine + e.StackTrace);
			} catch (Exception e) {
				AddinLogger.Error(e);
				OnUnknownError("DaisyUnexpectedError", e.GetType() + ": " + e.Message + " (" + e.StackTrace + ")");

				if (File.Exists(outputFile)) {
					File.Delete(outputFile);
				}

				return ConversionResult.UnknownError(e.Message + Environment.NewLine + e.StackTrace);
			} finally {
				DeleteTemporaryImages();
			}

			return ConversionResult.Success();
		}
        #endregion

        #region help methods

        private void ExecuteScript(string inputDaisyXmlPath)
		{
			if (ScriptToExecute != null)
			{
				try
				{
					ScriptToExecute.ExecuteScript(inputDaisyXmlPath);
				}
				catch (Exception e)
				{
					OnUnknownError(e.Message);
				}
			}
		}

		private void DeleteDTBookFilesIfExists(string outputfilepath)
		{
			if (File.Exists(outputfilepath + "\\dtbookbasic.css"))
				File.Delete(outputfilepath + "\\dtbookbasic.css");
			if (File.Exists(outputfilepath + "\\dtbook-2005-3.dtd"))
				File.Delete(outputfilepath + "\\dtbook-2005-3.dtd");
		}

		private string DocumentSkipped(ArrayList notTranslatedDoc)
		{
			string message = "";
			if (notTranslatedDoc.Count != 0)
			{
				for (int i = 0; i < notTranslatedDoc.Count; i++)
					message = message + Convert.ToString(i + 1) + ". " + notTranslatedDoc[i].ToString() + "\n";

				message = "\n\n" + "Files which are not in Word 2007 format are skipped during Translation:" + "\n" + message;
			}
			return message;
		}

		private void DeleteTemporaryImages()
		{
			string[] files = Directory.GetFiles(ConverterHelper.AppDataSaveAsDAISYDirectory);
			foreach (string file in files)
			{
				if (file.Contains(".jpg") || file.Contains(".JPG") || file.Contains(".PNG") || file.Contains(".png"))
				{
					File.Delete(file);
				}
			}
		}

		#endregion

		#region Multiple OOXML

		/// <summary>
		/// Function to check Whether Docs are already open or not
		/// </summary>
		/// <param name="listSubDocs">List of Sub Documents</param>
		/// <returns>Message whether any doc in List are open or not</returns>
		public static string CheckFileOPen(ArrayList listSubDocs)
		{
			String resultSubDoc = "notopen";
			for (int i = 0; i < listSubDocs.Count; i++)
			{
				string[] splt = listSubDocs[i].ToString().Split('|');
				try
				{
					Package pack;
					pack = Package.Open(splt[0].ToString(), FileMode.Open, FileAccess.ReadWrite);
					pack.Close();
				}
				catch (Exception e)
				{
					AddinLogger.Error(e);
					resultSubDoc = "open";
				}
			}
			return resultSubDoc;
		}

		#endregion

		#region virtual methods

		protected virtual void OnSuccessMasterSubValidation(string message)
		{

		}

		protected virtual void OnUnknownError(string error)
		{

		}

		protected virtual void OnUnknownError(string title, string details)
		{
		}

		protected virtual void OnValidationError(string error, string inputFile, string outputFile)
		{
		}

		protected virtual void OnLostElements(string inputFile, string outputFile, ArrayList elements)
		{
		}

		protected virtual bool IsContinueDTBookGenerationOnLostElements()
		{
			return true;
		}

		protected virtual void OnSuccess()
		{
		}

		protected virtual void OnMasterSubValidationError(string error)
		{

		}

		#endregion
	}
}