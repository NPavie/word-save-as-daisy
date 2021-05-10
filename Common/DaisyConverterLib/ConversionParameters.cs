using System;
using System.Collections;
using System.Collections.Generic;

namespace Daisy.DaisyConverter.DaisyConverterLib
{
    /// <summary>
    /// Input parameters for convert operations.
    /// This class also include the former TranslationParametersBuilder
    /// </summary>
    public class ConversionParameters
    {
        public string InputFile { get; set; }
        public string TempInputFile { get; set; }
        public string TempInputA { get; set; }
        
        public string ControlName { get; set; }
        public ArrayList ObjectShapes { get; set; }
        public Hashtable ListMathMl { get; set; }
        public ArrayList ImageIds { get; set; }
        public ArrayList InlineShapes { get; set; }
        public ArrayList InlineIds { get; set; }
        
		public string ScriptPath { get; set; }
		public string Directory { get; set; }

        public string TempOutputFile { get; set; }

        public string PipelineOutput { get; set; }

        // From the "TranslationParametersBuilder" class
        public string OutputFile { get; set; }
        public string Title { get; set; }
        public string Creator { get; set; }
        public string Publisher { get; set; }
        public string UID { get; set; }
        public string TrackChanges { get; set; }
        public string Subject { get; set; }
        public string Version { get; set; }
        public string ParseSubDocuments { get; set; }

        public string GetInputFileNameWithoutExtension
        {
            get
            {
                int lastSeparatorIndex = InputFile.LastIndexOf('\\');
                // Special case : onedrive documents uses https based URL format with '/' as separator
                if(lastSeparatorIndex < 0) {
                    lastSeparatorIndex = InputFile.LastIndexOf('/');
                }
                if (lastSeparatorIndex < 0) { // no path separator found
                    return InputFile.Remove(InputFile.LastIndexOf('.'));
                } else {
                    string tempInput = InputFile.Substring(lastSeparatorIndex);
                    return tempInput.Remove(tempInput.LastIndexOf('.'));
                }
            }
        }

        /// <summary>
        /// Function to mimic the TranslationParametersBuilder with* construction
        /// </summary>
        /// <param name="name">Name of the Class field to set</param>
        /// <param name="value">Object to assign to the field (this object will type casted to the targeted parameter type) </param>
        /// <returns>The converter itself</returns>
        public ConversionParameters withParameter(string name, object value) {
            switch (name) {
                case "InputFile":
                    InputFile = (string)value; break;
                case "TempInputFile":
                    TempInputFile = (string)value; break;
                case "TempInputA":
                    TempInputA = (string)value; break;
                case "ControlName":
                    ControlName = (string)value; break;
                case "ObjectShapes":
                    ObjectShapes = (ArrayList)value; break;
                case "ListMathMl":
                    ListMathMl = (Hashtable)value; break;
                case "ImageIds":
                    ImageIds = (ArrayList)value; break;
                case "InlineShapes":
                    InlineShapes = (ArrayList)value; break;
                case "InlineIds":
                    InlineIds = (ArrayList)value; break;
                case "ScriptPath":
                    ScriptPath = (string)value; break;
                case "Directory":
                    Directory = (string)value; break;
                case "OutputFile":
                    OutputFile = (string)value; break;
                case "Title":
                    Title = (string)value; break;
                case "Creator":
                    Creator = (string)value; break;
                case "Publisher":
                    Publisher = (string)value; break;
                case "UID":
                    UID = (string)value; break;
                case "TrackChanges":
                    TrackChanges = (string)value; break;
                case "Subject":
                    Subject = (string)value; break;
                case "Version":
                    Version = (string)value; break;
                case "ParseSubDocuments":
                    ParseSubDocuments = (string)value; break;
                default:
                    break;
            }
            return this;
        }

        /// <summary>
        /// Get the conversion settings hashtable (to replace the TranslationParametersBuilder behavior)
        /// </summary>
        public Hashtable ConversionParametersHash {
            get {
                Hashtable parameters = new Hashtable();
                
                if (OutputFile != null) parameters.Add("OutputFile", OutputFile);
                if (Title != null) parameters.Add("Title", Title);
                if (Creator != null) parameters.Add("Creator", Creator);
                if (Publisher != null) parameters.Add("Publisher", Publisher);
                if (UID != null) parameters.Add("UID", UID);
                if (Subject != null) parameters.Add("Subject", Subject);
                if (Version != null) parameters.Add("Version", Version);
                // TO BE CHANGED if the value changes in xslts
                if (TrackChanges != null) parameters.Add("TRACK", TrackChanges);
                if (ParseSubDocuments != null) parameters.Add("MasterSubFlag", ParseSubDocuments);

                // also retrieve global settings
                ConverterSettings globalSettings = new ConverterSettings();
                string imgoption = globalSettings.GetImageOption;
                string resampleValue = globalSettings.GetResampleValue;
                string characterStyle = globalSettings.GetCharacterStyle;
                string pagenumStyle = globalSettings.GetPagenumStyle;
                if (imgoption != " ") {
                    parameters.Add("ImageSizeOption", imgoption);
                    parameters.Add("DPI", resampleValue);
                }
                if (characterStyle != " ") {
                    parameters.Add("CharacterStyles", characterStyle);
                }
                if (pagenumStyle != " ") {
                    parameters.Add("Custom", pagenumStyle);
                }

                return parameters;
            }
        }
    }
}