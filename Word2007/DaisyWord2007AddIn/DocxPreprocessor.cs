using Daisy.SaveAsDAISY.Conversion;
using Daisy.SaveAsDAISY.Conversion.Events;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DaisyWord2007AddIn {
    public class DocxPreprocessor : IDocumentPreprocessor {

        object missing = Type.Missing;
        // Duplicate the current doc and use the copy
        object addToRecentFiles = false;
        object readOnly = false;

        // visibility
        object visible = true;
        object invisible = false;
        object originalFormat = WdOriginalFormat.wdOriginalDocumentFormat;
        object format = WdSaveFormat.wdFormatXMLDocument;

        Microsoft.Office.Interop.Word.Application currentInstance;
        public DocxPreprocessor(Microsoft.Office.Interop.Word.Application WordInstance) {
            currentInstance = WordInstance;
        }

        public PreProcessingStatus CreateWorkingCopy(ref object preprocessedObject, DocumentParameters document, IConversionEventsHandler eventsHandler = null) {
            Microsoft.Office.Interop.Word.Document currentDoc = (Microsoft.Office.Interop.Word.Document)preprocessedObject;
            object tmpFileName = document.CopyPath;
            object originalPath = document.InputPath;
            // Save a copy and reopen the the original document
            currentDoc.SaveAs(ref tmpFileName, ref format, ref missing, ref missing, ref addToRecentFiles, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            currentDoc.Close();

            // Open, or retrieve the temp file if opened in word
            Document newDoc = currentInstance.Documents.Open(ref tmpFileName, ref missing, ref readOnly, ref addToRecentFiles, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref invisible, ref missing, ref missing, ref missing, ref missing);
            // close the temp file
            object saveChanges = WdSaveOptions.wdDoNotSaveChanges;

            // Close the new doc and reopen the original one
            newDoc.Close(ref saveChanges, ref originalFormat, ref missing);
            currentDoc = currentInstance.Documents.Open(ref originalPath);

            return PreProcessingStatus.CreatedWorkingCopy;

        }

        public PreProcessingStatus endPreprocessing(ref object preprocessedObject, IConversionEventsHandler eventsHandler = null) {
            throw new NotImplementedException();
        }

        public PreProcessingStatus ProcessEquations(ref object preprocessedObject, DocumentParameters document, IConversionEventsHandler eventsHandler = null) {
            throw new NotImplementedException();
        }

        public PreProcessingStatus ProcessInlineShapes(ref object preprocessedObject, DocumentParameters document, IConversionEventsHandler eventsHandler = null) {
            throw new NotImplementedException();
        }

        public PreProcessingStatus ProcessShapes(ref object preprocessedObject, DocumentParameters document, IConversionEventsHandler eventsHandler = null) {
            throw new NotImplementedException();
        }

        public object startPreprocessing(DocumentParameters document, IConversionEventsHandler eventsHandler = null) {
            object path = (object)document.InputPath;
            Microsoft.Office.Interop.Word.Document currentDoc = currentInstance.Documents.Open(ref path);

            return currentDoc;
        }

        public PreProcessingStatus ValidateName(ref object preprocessedObject, FilenameValidator authorizedNamePattern, IConversionEventsHandler eventsHandler = null) {
            Microsoft.Office.Interop.Word.Document currentDoc = (Microsoft.Office.Interop.Word.Document)preprocessedObject;
            Microsoft.Office.Interop.Word.Application WordInstance = currentDoc.Application;
            bool nameIsValid = false;
            do {
                bool docIsRenamed = false;
                if (!authorizedNamePattern.AuthorisationPattern.IsMatch(currentDoc.Name)) { // check only name (i assume it may still lead to problem if path has commas)
                    
                    DialogResult userAnswer = MessageBox.Show(authorizedNamePattern.unauthorizedNameMessage, "Unauthorized characters in the document filename", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                    if (userAnswer == DialogResult.Yes) {
                        Dialog dlg = WordInstance.Dialogs[WdWordDialog.wdDialogFileSaveAs];
                        int saveResult = dlg.Show(ref missing);
                        if (saveResult == -1) { // ok pressed, see https://docs.microsoft.com/fr-fr/dotnet/api/microsoft.office.interop.word.dialog.show?view=word-pia#Microsoft_Office_Interop_Word_Dialog_Show_System_Object__
                            docIsRenamed = true;
                        } else return PreProcessingStatus.Canceled;// PreprocessingData.Canceled("User canceled a renaming request for an invalid docx filename");
                    } else if (userAnswer == DialogResult.Cancel) {
                        return PreProcessingStatus.Canceled;// PreprocessingData.Canceled("User canceled a renaming request for an invalid docx filename");
                    }
                    // else a sanitize path in the DaisyAddinLib will replace commas by underscore.
                    // Other illegal characters regarding the conversion to DAISY book are replaced by underscore by the pipeline itself
                    // While image names seems to be sanitized in other process
                }
                nameIsValid = !docIsRenamed;
            } while (!nameIsValid);
            return PreProcessingStatus.ValidatedName;
        }
    }
}
