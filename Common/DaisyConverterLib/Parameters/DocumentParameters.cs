using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Daisy.SaveAsDAISY.Conversion {

    /// <summary>
    /// Document specific parameters, to be extracted by Preprocessors
    /// </summary>
    public class DocumentParameters {

        public enum DocType {
            Simple,
            Master,
            Sub
        }


        


        public DocumentParameters(string inputPath) {
            this.InputPath = inputPath;
            
            ListMathMl = new Hashtable();
            ObjectShapes = new List<string>();
            ImageIds = new List<string>();
            InlineShapes = new List<string>();
            InlineIds = new List<string>();
        }

        public PreProcessingStatus preprocessingStatus = PreProcessingStatus.None;


        public DocumentParameters prepareForConversion(
                IDocumentPreprocessor documentPreprocessor,
                FilenameValidator authorizedFilenameFormat,
                Events.IConversionEventsHandler eventsHandler) {
            
            this.CopyPath = ConverterHelper.GetTempPath(InputPath, ".docx");
            object preprocessedObject = documentPreprocessor.startPreprocessing(this);
            try {
                do {
                    switch (this.preprocessingStatus) {
                        case PreProcessingStatus.None: // Starting by validating file name
                            this.preprocessingStatus = documentPreprocessor.ValidateName(ref preprocessedObject, authorizedFilenameFormat, eventsHandler);
                            break;
                        case PreProcessingStatus.ValidatedName: // make the working copy
                            this.preprocessingStatus = documentPreprocessor.CreateWorkingCopy(ref preprocessedObject, this.CopyPath, eventsHandler);
                            break;
                        case PreProcessingStatus.CreatedWorkingCopy: // start processing shapes
                            this.preprocessingStatus = documentPreprocessor.ProcessShapes(ref preprocessedObject, this, eventsHandler);
                            break;
                        case PreProcessingStatus.ProcessedShapes: // start processing inline shapes
                            this.preprocessingStatus = documentPreprocessor.ProcessInlineShapes(ref preprocessedObject, this, eventsHandler);
                            break;
                        case PreProcessingStatus.ProcessedInlineShapes: // start processing math
                            this.preprocessingStatus = documentPreprocessor.ProcessEquations(ref preprocessedObject, this, eventsHandler);
                            break;
                        case PreProcessingStatus.ProcessedMathML: // finaliz
                            this.preprocessingStatus = documentPreprocessor.endPreprocessing(ref preprocessedObject, eventsHandler);
                            break;
                    }
                } while (this.preprocessingStatus != PreProcessingStatus.Canceled &&
                     this.preprocessingStatus != PreProcessingStatus.Error &&
                     this.preprocessingStatus != PreProcessingStatus.Success);
            } catch (Exception e) {
                eventsHandler.onPreprocessingError(this, e.Message);
            } finally {
                if(this.preprocessingStatus != PreProcessingStatus.Success) {
                    documentPreprocessor.endPreprocessing(ref preprocessedObject, eventsHandler);
                }
            }

            SubdocumentsList subdocuments = SubdocumentsManager.FindSubdocuments(
                InputPath,
                CopyPath);
            if(subdocuments.Errors.Count > 0) {
                string errors = "The following errors where found searching for subdocuments:\r\n" + string.Join("\r\n", subdocuments.Errors);
                this.preprocessingStatus = PreProcessingStatus.Error;
                eventsHandler.onPreprocessingError(this, errors);
            } else if (!subdocuments.Empty) {
                
            }

            return this;
        }


        /// <summary>
        /// Word document type between :<br/>
        /// - Simple : the document is self contained<br/>
        /// - Master : the document refers to subdocuments<br/>
        /// - Sub : the document is refered by another document <br/>
        /// Note : <br/>
        /// - Master Document will have subdocuments <br/>
        /// - SubDocument will have a resource ID <br/>
        /// - Simple will have neither
        /// </summary>
        public DocType Type {
            get {
                if(this.ResourceId != null) {
                    return DocType.Sub;
                } else if (this.SubDocuments.Count > 0) {
                    return DocType.Master;
                } else return  DocType.Simple;
            }
        }

        /// <summary>
        /// Original path/URL of the input
        /// </summary>
        public string InputPath { get; set; }

        /// <summary>
        /// Document copy to use for processing
        /// </summary>
        public string CopyPath { get; set; }

        /// <summary>
        ///
        /// </summary>
        public string OutputPath { get; set; }


        /// <summary>
        /// 
        /// </summary>
        //public DocType Type { get; set; }

        public List<string> ObjectShapes { get; set; }

        /// <summary>
        /// To be replaced by a list of string later on
        /// (Needs to update the DaisyClass object for that)
        /// </summary>
        public Hashtable ListMathMl { get; set; }
        public List<string> ImageIds { get; set; }
        public List<string> InlineShapes { get; set; }
        public List<string> InlineIds { get; set; }

        

        /// <summary>
        /// Sub documents referenced by the current document
        /// </summary>
        public List<DocumentParameters> SubDocuments { get; set; }

        /// <summary>
        /// Resource ID of the document it it is a sub document contained in a Master document
        /// </summary>
        public string ResourceId { get; set; }

        public string GetInputFileNameWithoutExtension {
            get {
                int lastSeparatorIndex = InputPath.LastIndexOf('\\');
                // Special case : onedrive documents uses https based URL format with '/' as separator
                if (lastSeparatorIndex < 0) {
                    lastSeparatorIndex = InputPath.LastIndexOf('/');
                }
                if (lastSeparatorIndex < 0) { // no path separator found
                    return InputPath.Remove(InputPath.LastIndexOf('.'));
                } else {
                    string tempInput = InputPath.Substring(lastSeparatorIndex);
                    return tempInput.Remove(tempInput.LastIndexOf('.'));
                }
            }
        }
    }
}
