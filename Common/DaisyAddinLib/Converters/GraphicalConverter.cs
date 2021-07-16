using System;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;

using Daisy.SaveAsDAISY.DaisyConverterLib;

namespace Daisy.SaveAsDAISY
{
	/// <summary>
	/// Implements convertion with UI messages/dialogs.
    /// - Request user confirmation p
    /// - Progression dialog : display progression message in a textarea instead of progress bar
    /// 
	/// </summary>
    public class GraphicalConverter : BaseConverter
    {
        ConversionProgress progressDialog;

        bool progressDialogOpened = false;

        public GraphicalConverter(WordToDTBookXMLTransform converter) 
			: base(converter, null)
        {
            progressDialog = new ConversionProgress();
            progressDialog.setCancelClickListener(this.onConversionCancel);
        }

        public GraphicalConverter(WordToDTBookXMLTransform converter, ScriptParser scriptToExecute) 
			: base(converter, scriptToExecute)
        {
        }


        #region Overrides of AbstractConverter

        protected override void OnValidationError(string error, string inputFile, string outputFile)
        {
            Validation validationDialog = new Validation("FailedLabel", error, inputFile, outputFile, ResManager);
            validationDialog.ShowDialog();
        }

        protected override bool IsContinueDTBookGenerationOnLostElements()
        {
            DialogResult continueDTBookGenerationResult = MessageBox.Show("Do you want to create audio file", "SaveAsDAISY", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            return continueDTBookGenerationResult == DialogResult.Yes;
        }

        protected override void OnSuccess()
        {
            MessageBox.Show(ResManager.GetString("SucessLabel"), "SaveAsDAISY - Success", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        }

        protected override void OnMasterSubValidationError(string error)
        {
            MasterSubValidation infoBox = new MasterSubValidation(error, "Validation");
            infoBox.ShowDialog();
        }

        protected override void OnSuccessMasterSubValidation(string message)
        {
            MasterSubValidation infoBox = new MasterSubValidation(message, "Success");
            infoBox.ShowDialog();
        }

        protected override void OnUnknownError(string error)
        {
            if (progressDialogOpened) {
                progressDialog.Close();
                progressDialogOpened = false;
            }
            MessageBox.Show(error, "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        protected override void OnUnknownError(string title, string details)
        {
            if (progressDialogOpened) {
                progressDialog.Close();
                progressDialogOpened = false;
            }
            InfoBox infoBox = new InfoBox(title, details, ResManager);
            infoBox.ShowDialog();
        }

        protected override void onDocumentConversionStart(DocumentParameters document, ConversionParameters conversion) {
            base.onDocumentConversionStart(document, conversion);
            if (!progressDialogOpened) {
                progressDialog.Show();
                progressDialogOpened = true;
            }
            progressDialog.addMessage("Converting document " + document.InputPath);
        }

        protected override void onProgressMessageReceived(object sender, EventArgs e) {
            base.onProgressMessageReceived(sender, e);
            progressDialog.addMessage(((DaisyEventArgs)e).Message);
        }

        protected override void onFeedbackMessageReceived(object sender, EventArgs e) {
            base.onFeedbackMessageReceived(sender, e);
            progressDialog.addMessage(((DaisyEventArgs)e).Message);
        }

        protected override void onFeedbackValidationMessageReceived(object sender, EventArgs e) {
            base.onFeedbackValidationMessageReceived(sender, e);
            progressDialog.addMessage(((DaisyEventArgs)e).Message);

        }

        protected override void OnLostElements(string outputFile, ArrayList elements) {
            Fidility fidilityDialog = new Fidility("FeedbackLabel", elements, outputFile, ResManager);
            fidilityDialog.ShowDialog();
            base.OnLostElements(outputFile, elements);
        }

        protected override void onDocumentListConversionStart(List<DocumentParameters> documentLists, ConversionParameters conversion) {
            base.onDocumentListConversionStart(documentLists, conversion);
            if (!progressDialogOpened) {
                progressDialog.Show();
                progressDialogOpened = true;
            }
        }

        protected override void onPostProcessingStart(ConversionParameters conversion) {
            base.onPostProcessingStart(conversion);
        }

        protected override void onDocumentListConversionSuccess(List<DocumentParameters> documentLists, ConversionParameters conversion) {
            base.onDocumentListConversionSuccess(documentLists, conversion);
            // close the progression dialog
            if (progressDialogOpened) {
                progressDialog.Close();
                progressDialogOpened = false;
            }
        }

        protected override void onDocumentConversionSuccess(DocumentParameters document, ConversionParameters conversion) {
            base.onDocumentConversionSuccess(document, conversion);
            // close the progression dialog
            if (progressDialogOpened) {
                progressDialog.Close();
                progressDialogOpened = false;
            }
            // Not sure if necessary but just in case the close function would raised the formClosing event
            this.isCanceled = false;
        }

        protected override void onPostProcessingSuccess(ConversionParameters conversion) {
            base.onPostProcessingSuccess(conversion);
        }

        protected override void onConversionCancel() {
            base.onConversionCancel();
            // close the progression dialog
            if (progressDialogOpened) {
                progressDialog.Close();
                progressDialogOpened = false;
            }
        }


        #endregion
    }
}