using Daisy.SaveAsDAISY.Conversion;
using Daisy.SaveAsDAISY.Conversion.Events;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace Daisy.SaveAsDAISY {
    /// <summary>
    /// Conversion events handler using graphical interface (windows dialog)
    /// </summary>
	public class GraphicalEventsHandler : IConversionEventsHandler
	{
        private System.Resources.ResourceManager resourceManager = new System.Resources.ResourceManager(
                    "DaisyAddinLib.resources.Labels",
                    Assembly.GetExecutingAssembly()
            );

        private ConversionProgress progressDialog;

        public GraphicalEventsHandler() {

        }

        public void LinkToProgressDialog(ref ConversionProgress progressDialog) {
            this.progressDialog = progressDialog;
        }

        public System.Resources.ResourceManager ResManager {
            get { return this.resourceManager; }
        }
        public void OnStop(string message)
		{
			OnStop(message,"SaveAsDAISY");
		}

		public bool AskForTranslatingSubdocuments()
		{
			DialogResult dialogResult = MessageBox.Show("Do you want to translate the current document along with sub documents?", "SaveAsDAISY", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
			return dialogResult == DialogResult.Yes;
		}

		public void OnError(string errorMessage)
		{
            if (progressDialog != null && !progressDialog.IsDisposed) {
                progressDialog.Close();
            }
            MessageBox.Show(errorMessage, "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Error);
		}

		public void OnStop(string message, string title)
		{
            if (progressDialog != null && !progressDialog.IsDisposed) {
                progressDialog.Close();
            }
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Stop);
		}

        public void onDocumentListConversionStart(List<DocumentParameters> documentLists, ConversionParameters conversion) {
            if (progressDialog != null && !progressDialog.IsDisposed) {
                if (!progressDialog.Visible) progressDialog.Show();
                progressDialog.addMessage("Starting documents list conversion");
            }
        }

        public void onDocumentConversionStart(DocumentParameters document, ConversionParameters conversion) {
            if (progressDialog != null && !progressDialog.IsDisposed) {
                if(!progressDialog.Visible) progressDialog.Show();
                progressDialog.addMessage("Converting document " + document.InputPath);
            }
            
        }

        public void onPostProcessingStart(ConversionParameters conversion) {
            if (progressDialog != null && !progressDialog.IsDisposed) {
                progressDialog.Close();
            }
        }

        public void onDocumentListConversionSuccess(List<DocumentParameters> documentLists, ConversionParameters conversion) {
            if (progressDialog != null && !progressDialog.IsDisposed) {
                progressDialog.Close();
            }
        }

        public void onDocumentConversionSuccess(DocumentParameters document, ConversionParameters conversion) {
            if (progressDialog != null && !progressDialog.IsDisposed) {
                if (!progressDialog.Visible) progressDialog.Show();
                progressDialog.addMessage("-> Success");
            }
        }

        public void onPostProcessingSuccess(ConversionParameters conversion) {
            
        }

        public void onConversionCanceled() {
            if (progressDialog != null && !progressDialog.IsDisposed) {
                progressDialog.Close();
            }
        }

        public void onProgressMessageReceived(object sender, EventArgs e) {
            if (progressDialog != null && !progressDialog.IsDisposed) {
                if (!progressDialog.Visible) progressDialog.Show();
                progressDialog.addMessage(((DaisyEventArgs)e).Message);
            }
        }

        public void onFeedbackMessageReceived(object sender, EventArgs e) {
            if (progressDialog != null && !progressDialog.IsDisposed) {
                if (!progressDialog.Visible) progressDialog.Show();
                progressDialog.addMessage(((DaisyEventArgs)e).Message);
            }
        }

        public void onFeedbackValidationMessageReceived(object sender, EventArgs e) {
            if (progressDialog != null && !progressDialog.IsDisposed) {
                if (!progressDialog.Visible) progressDialog.Show();
                progressDialog.addMessage(((DaisyEventArgs)e).Message);
            }
        }

        public void OnSuccessMasterSubValidation(string message) {
            MasterSubValidation infoBox = new MasterSubValidation(message, "Success");
            infoBox.ShowDialog();
        }

        public void OnUnknownError(string error) {
            if (progressDialog != null && !progressDialog.IsDisposed) {
                progressDialog.Close();
            }
            MessageBox.Show(error, "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public void OnUnknownError(string title, string details) {
            if (progressDialog != null && !progressDialog.IsDisposed) {
                progressDialog.Close();
            }
            InfoBox infoBox = new InfoBox(title, details, ResManager);
            infoBox.ShowDialog();
        }

        public void OnValidationErrors(List<ValidationError> errors, string outputFile) {
            Validation validationDialog = new Validation(
                "FailedLabel",
                string.Join(
                        "\r\n",
                        errors.Select(
                            error => error.ToString()
                        ).ToArray()
                    ), outputFile, 
                ResManager);
            validationDialog.ShowDialog();
        }

        public void OnLostElements(string outputFile, ArrayList elements) {
            Fidility fidilityDialog = new Fidility("FeedbackLabel", elements, outputFile, ResManager);
            fidilityDialog.ShowDialog();
        }

        public bool IsContinueDTBookGenerationOnLostElements() {
            DialogResult continueDTBookGenerationResult = MessageBox.Show("Do you want to create audio file", "SaveAsDAISY", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            return continueDTBookGenerationResult == DialogResult.Yes;
        }

        public void OnSuccess() {
            if (progressDialog != null && !progressDialog.IsDisposed) {
                progressDialog.Close();
            }
            MessageBox.Show(ResManager.GetString("SucessLabel"), "SaveAsDAISY - Success", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        }

        public void OnMasterSubValidationError(string error) {
            MasterSubValidation infoBox = new MasterSubValidation(error, "Validation");
            infoBox.ShowDialog();
        }

        public void onDocumentPreprocessingStart(DocumentParameters document) {
            throw new NotImplementedException();
        }

        public void onPreprocessingCancel() {
            throw new NotImplementedException();
        }

        public void onPreprocessingError(DocumentParameters document, string errors) {
            throw new NotImplementedException();
        }

        public void onPreprocessingSuccess() {
            throw new NotImplementedException();
        }
    }
}