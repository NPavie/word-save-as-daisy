/**
 * Preprocessing of the word document for the 2007 addin
 */

using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using MSword = Microsoft.Office.Interop.Word;

using IConnectDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;
using ConnectFORMATETC = System.Runtime.InteropServices.ComTypes.FORMATETC;
using ConnectSTGMEDIUM = System.Runtime.InteropServices.ComTypes.STGMEDIUM;
using COMException = System.Runtime.InteropServices.COMException;
using TYMED = System.Runtime.InteropServices.ComTypes.TYMED;
using System.Collections;

using Daisy.SaveAsDAISY.Conversion;
using System.Drawing.Imaging;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using System.IO.Packaging;
using System.Threading;
using System.Xml;
using Daisy.SaveAsDAISY;
using Daisy.SaveAsDAISY.Conversion.Events;

namespace DaisyWord2007AddIn {



    /// <summary>
    /// MS Word processing functions meant to be used in the addin or external C# programs with word interop capabilities
    /// </summary>
    public static class WordPreprocessing {

        /// <summary>
        /// Function to generate Random ID
        /// </summary>
        /// <returns></returns>
        static public long GenerateId() {
            byte[] buffer = Guid.NewGuid().ToByteArray();
            return BitConverter.ToInt64(buffer, 0);
        }

        

        public static PreprocessingData prepareConversion(
            IConversionEventsHandler eventsHandler,
            MSword.Document documentToConvert,
            Addin addin = null,
            string conversionMode = "DaisySingle",
            Pipeline postprocessingPipeline = null
        ) {
            MSword.Document currentDoc = documentToConvert;
            MSword.Application WordInstance = documentToConvert.Application;
            PreprocessingData result = new PreprocessingData(
                WordInstance.Version, 
                postprocessingPipeline, 
                AddInHelper.buttonIsSingleWordToXMLConversion(conversionMode) ? "" : conversionMode); // Use default postprocessing script
            int fileIndex;

            if (addin == null) {
                addin = new Daisy.SaveAsDAISY.Addin();
            }

            if (!currentDoc.Saved || currentDoc.FullName.LastIndexOf('.') < 0) {
                eventsHandler.OnStop(addin.GetString("DaisySaveDocumentBeforeExport"));
                return PreprocessingData.Failed("Please save your document before going further.");
            }

            fileIndex = currentDoc.FullName.LastIndexOf('.');
            String substr = currentDoc.FullName.Substring(fileIndex);

            if (substr.ToLower() != ".docx") {
                eventsHandler.OnStop(addin.GetString("DaisySaveDocumentin2007"));
                return PreprocessingData.Failed("The document is not a docx file saved on your system.");
            }
            object missing = Type.Missing;

            // Adding a filename check on the current docx to prevent problematic characters in the filename
            StringBuilder errorFileNameMessage = new StringBuilder("Your document file name contains unauthorized characters, that may be automatically replaced by underscores.\r\n");
            // For dtbook conversion, any correct system file name will work, except for the ones with commas in it in my tests
            // Possibly an error in the pipeline commande line parsing of arguments in the pipeline side
            string authorizedNamePattern = @"^[^,]+$";
            if (AddInHelper.buttonIsSingleWordToXMLConversion(conversionMode)) {
                errorFileNameMessage.Append("Any commas (,) present in the file name should be removed, or they will be replaced by underscores automatically.");
            } else {
                // TODO : specific name pattern following daisy book naming convention to find
                authorizedNamePattern = @"^[a-zA-Z0-9_\-\.]+\.docx$";
                errorFileNameMessage.Append(
                    "Only Alphanumerical letters (a-z, A-Z, 0-9), hyphens (-), dots (.) " +
                        "and underscores (_) are allowed in DAISY file names." +
                    "\r\nAny other characters (including spaces) will be replaced automaticaly by underscores.");
            }
            errorFileNameMessage.Append(
                "\r\n" +
                "\r\nDo you want to save this document under a new name ?" +
                "\r\nThe document with the original name will not be deleted." +
                "\r\n" +
                "\r\n(Click Yes to save the document under a new name and use the new one, " +
                    "No to continue with the current document, " +
                    "or Cancel to abort the conversion)");


            Regex validator = new Regex(authorizedNamePattern);
            bool nameIsValid;
            do {
                bool docIsRenamed = false;
                if (!validator.IsMatch(currentDoc.Name)) { // check only name (i assume it may still lead to problem if path has commas)
                    DialogResult userAnswer = MessageBox.Show(errorFileNameMessage.ToString(), "Unauthorized characters in the document filename", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                    if (userAnswer == DialogResult.Yes) {
                        Dialog dlg = WordInstance.Dialogs[MSword.WdWordDialog.wdDialogFileSaveAs];
                        int saveResult = dlg.Show(ref missing);
                        if (saveResult == -1) { // ok pressed, see https://docs.microsoft.com/fr-fr/dotnet/api/microsoft.office.interop.word.dialog.show?view=word-pia#Microsoft_Office_Interop_Word_Dialog_Show_System_Object__
                            docIsRenamed = true;
                        } else return PreprocessingData.Canceled("User canceled a renaming request for an invalid docx filename");
                    } else if (userAnswer == DialogResult.Cancel) {
                        return PreprocessingData.Canceled("User canceled a renaming request for an invalid docx filename");
                    }
                    // else a sanitize path in the DaisyAddinLib will replace commas by underscore.
                    // Other illegal characters regarding the conversion to DAISY book are replaced by underscore by the pipeline itself
                    // While image names seems to be sanitized in other process
                }
                nameIsValid = !docIsRenamed;
            } while (!nameIsValid);

            object originalPath = currentDoc.FullName;
            object tmpFileName = addin.GetTempPath((string)originalPath, ".docx");
            //object newName = Path.GetTempFileName() + Path.GetExtension((string)originalPath);

            // Duplicate the current doc and use the copy
            object addToRecentFiles = false;
            object readOnly = false;

            // visibility
            object visible = true;
            object invisible = false;
            object originalFormat = WdOriginalFormat.wdOriginalDocumentFormat;
            object format = WdSaveFormat.wdFormatXMLDocument;

            // FIX 05/03/2021 : Error is raised here for onedrive files that are using "http(s)" urls
            // For now we replace the copy by a standard office save and reopen the original file

            // Save a copy and reopen the the original document
            currentDoc.SaveAs(ref tmpFileName, ref format, ref missing, ref missing, ref addToRecentFiles, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            currentDoc.Close();

            // Open, or retrieve the temp file if opened in word
            Document newDoc = WordInstance.Documents.Open(ref tmpFileName, ref missing, ref readOnly, ref addToRecentFiles, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref invisible, ref missing, ref missing, ref missing, ref missing);
            // close the temp file
            object saveChanges = WdSaveOptions.wdDoNotSaveChanges;

            // Close the new doc and reopen the original one
            newDoc.Close(ref saveChanges, ref originalFormat, ref missing);
            currentDoc = WordInstance.Documents.Open(ref originalPath);

            string docFile = (string)tmpFileName;

            PrepopulateDaisyXml prepopulateDaisyXml = new PrepopulateDaisyXml(docFile);
            
            prepopulateDaisyXml.Save();

            // FIX 05/03/2021 : OriginalFilePath is used as InputFile in OoxToDaisyParameters. 
            // For onedrive url (starting with https) the temp copy is used as input path instead of the real one
            result.conversionParameters.InputFile = currentDoc.FullName.StartsWith("http") ? docFile : currentDoc.FullName;
            result.conversionParameters.TempInputFile = docFile;
            result.conversionParameters.ParseSubDocuments = PreprocessingData.requestSubDocumentsConversion(docFile, eventsHandler);
            //result.InitializeWindow.Show();
            //Application.DoEvents();
            try {
                Exception threadEx = null;
                Thread staThread = new Thread(
                    delegate () {
                        try {
                            WordPreprocessing.exportShapes(
                                currentDoc,
                                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + "SaveAsDAISY");
                        } catch (Exception ex) {
                            threadEx = ex;
                        }
                    });
                staThread.SetApartmentState(ApartmentState.STA);
                staThread.Start();
                staThread.Join();
                if (threadEx != null) {
                    throw threadEx;
                }
            } catch (Exception e) {
                eventsHandler.OnError("An error occured while preprocessing shapes and may prevent the rest of the conversion to success:" +
                    "\r\n- " + e.Message +
                    "\r\n" + e.StackTrace);
            }
            try {
                Exception threadEx = null;
                Thread staThread = new Thread(
                    delegate () {
                        try {
                            WordPreprocessing.exportInlineShapes(
                                currentDoc,
                                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + "SaveAsDAISY");
                        } catch (Exception ex) {
                            threadEx = ex;
                        }
                    });
                staThread.SetApartmentState(ApartmentState.STA);
                staThread.Start();
                staThread.Join();
                if (threadEx != null) {
                    throw threadEx;
                }
            } catch (Exception e) {
                eventsHandler.OnError("An error occured while preprocessing images and may prevent the rest of the conversion to success:" +
                    "\r\n- " + e.Message +
                    "\r\n" + e.StackTrace);
            }
            WordPreprocessing.parseEquations(
                    eventsHandler,
                    currentDoc,
                    result.conversionParameters.ListMathMl);

            //this.applicationObject.ActiveDocument.Save();
            if (result.conversionParameters.ParseSubDocuments == "Yes") {
                result.IsSuccess = preprocessSubDocuments(eventsHandler, result, currentDoc);
            } else {
                
                result.IsSuccess = true;

            }
            //result.InitializeWindow.Close();

            return result;
        }

        #region Shapes

        static public void exportShapes(
            MSword.Document doc,
            string outputPath
        ) {
            object missing = Type.Missing;
            
            List<string> warnings = new List<string>();
            String fileName = doc.Name.ToString().Replace(" ", "_");
            MSword.Application WordInstance = doc.Application;
            WordInstance.Activate();
            
            System.Diagnostics.Process objProcess = System.Diagnostics.Process.GetCurrentProcess();
            // If we need a finer way to find the process handle
            // https://stackoverflow.com/questions/8673726/get-specific-window-handle-using-office-interop

            foreach (MSword.Shape item in doc.Shapes) {
                if (!item.Name.Contains("Text Box")) {
                    item.Select(ref missing);
                    string pathShape = outputPath + "\\" + Path.GetFileNameWithoutExtension(fileName) + "-Shape" + item.ID.ToString() + ".png";
                    WordInstance.Selection.CopyAsPicture();
                    try {
                        // Note : using Clipboard.GetImage() set Word to display a clipboard data save request on closing
                        // So we rely on the user32 clipboard methods that does not seem to be intercepted by Office
                        System.Drawing.Image image = ClipboardEx.GetEMF(objProcess.MainWindowHandle);
                        byte[] Ret;
                        MemoryStream ms = new MemoryStream();
                        image.Save(ms, ImageFormat.Png);
                        Ret = ms.ToArray();
                        FileStream fs = new FileStream(pathShape, FileMode.Create, FileAccess.Write);
                        fs.Write(Ret, 0, Ret.Length);
                        fs.Flush();
                        fs.Dispose();
                    } catch (ClipboardDataException cde) {
                        warnings.Add("- Shape " + item.ID.ToString() + ": " + cde.Message);
                    } catch (Exception e) {
                        throw e;
                    } finally {
                        Clipboard.Clear();

                    }
                }
            }
            if (warnings.Count > 0) {
                string warningMessage = "Some shapes could not be exported from the document " + doc.Name;
                foreach (string warning in warnings) {
                    warningMessage += "\r\n" + warning;
                }
                throw new Exception(warningMessage);
            }
        }

        /// <summary>
        /// Save inline shapes from an opened document
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="outputFolderPath"></param>
        static public void exportInlineShapes(
            MSword.Document doc,
            string outputFolderPath
        ) {
            List<string> warnings = new List<string>();
            MSword.Application WordInstance = doc.Application;
            WordInstance.Activate();

            System.Diagnostics.Process objProcess = System.Diagnostics.Process.GetCurrentProcess();
            MSword.Range rng;
            String fileName = doc.Name.ToString().Replace(" ", "_");
            foreach (MSword.Range tmprng in doc.StoryRanges) {
                rng = tmprng;
                while (rng != null) {
                    foreach (MSword.InlineShape item in rng.InlineShapes) {
                        if ((item.Type.ToString() != "wdInlineShapeEmbeddedOLEObject") && ((item.Type.ToString() != "wdInlineShapePicture"))) {
                            string str = "Shapes_" + GenerateId().ToString();
                            string shapeOutputPath = Path.Combine(outputFolderPath, Path.GetFileNameWithoutExtension(fileName) + "-" + str + ".png");
                            object range = item.Range;
                            item.Range.Bookmarks.Add(str, ref range);
                            item.Range.CopyAsPicture();
                            try {
                                System.Drawing.Image image = ClipboardEx.GetEMF(objProcess.MainWindowHandle);
                                byte[] Ret;
                                MemoryStream ms = new MemoryStream();
                                image.Save(ms, ImageFormat.Png);
                                Ret = ms.ToArray();
                                FileStream fs = new FileStream(shapeOutputPath, FileMode.Create, FileAccess.Write);
                                fs.Write(Ret, 0, Ret.Length);
                                fs.Flush();
                                fs.Dispose();
                            } catch (ClipboardDataException cde) {
                                warnings.Add("- InlineShape with AltText \"" + item.AlternativeText.ToString() + "\": " + cde.Message);
                            } catch (Exception e) {
                                throw e;
                            } finally {
                                Clipboard.Clear();

                            }
                        }
                    }
                    rng = rng.NextStoryRange;
                }
            }

            if (warnings.Count > 0) {
                string warningMessage = "Some images could not be exported from the document " + doc.Name;
                foreach (string warning in warnings) {
                    warningMessage += "\r\n" + warning;
                }
                throw new Exception(warningMessage);
            }
        }

        #endregion

        #region MathML

        #region Imports from dll and COMs
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, ExactSpelling = true, SetLastError = true)]
        private static extern IntPtr GlobalLock(HandleRef handle);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, ExactSpelling = true, SetLastError = true)]
        private static extern bool GlobalUnlock(HandleRef handle);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, ExactSpelling = true, SetLastError = true)]
        private static extern int GlobalSize(HandleRef handle);

        [DllImport("ole32.dll", CharSet = CharSet.Auto, ExactSpelling = true, SetLastError = true)]
        private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

        [DllImport("ole32.dll", CharSet = CharSet.Auto, ExactSpelling = true, SetLastError = true)]
        private static extern int OleGetAutoConvert(ref Guid oCurrentCLSID, out Guid pConvertedClsid);

        [DllImport("ole32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool IsEqualGUID(ref Guid rclsid1, ref Guid rclsid);
        [ComImport]
        [Guid("00000112-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IOleObject {
            void DoNotCall_1(object x);
            void DoNotCall_2(ref object x);
            void SetHostNames(object szContainerApp, object szContainerObj);
            void Close(uint dwSaveOption);
        };

        public enum OLECLOSE {
            OLECLOSE_SAVEIFDIRTY = 0,
            OLECLOSE_NOSAVE = 1,
            OLECLOSE_PROMPTSAVE = 2,
        }
        #endregion

        /// <summary>
        /// Search the equations in the stories of a document.
        /// </summary>
        /// <param name="eventsHandler"></param>
        /// <param name="WordInstance">Instance of the word application</param>
        /// <param name="documentPath">Path of the document to parse</param>
        /// <param name="mathMLEquationsTable">Table to store the equations list per document story</param>
        /// <returns>The number of equations found</returns>
        static public int parseEquations(
            IConversionEventsHandler eventsHandler,
            MSword.Document doc,
            Hashtable mathMLEquationsTable
        ) {
            Int16 showMsg = 0;
            MSword.Range rng;
            String storyName = "";
            int iNumShapesViewed = 0;

            foreach (MSword.Range tmprng in doc.StoryRanges) {
                ArrayList listmathML = new ArrayList();
                rng = tmprng;
                storyName = rng.StoryType.ToString();
                while (rng != null) {
                    storyName = rng.StoryType.ToString();
                    MSword.InlineShapes shapes = rng.InlineShapes;
                    if (shapes != null && shapes.Count > 0) {
                        int iCount = 1;
                        int iNumShapes = 0;
                        Microsoft.Office.Interop.Word.InlineShape shape;
                        iNumShapes = shapes.Count;
                        //iCount is the LCV and the shapes accessor is 1 based, more that likely from VBA.

                        while (iCount <= iNumShapes) {
                            if (shapes[iCount].Type.ToString() == "wdInlineShapeEmbeddedOLEObject") {
                                if (shapes[iCount].OLEFormat.ProgID == "Equation.DSMT4") {
                                    shape = shapes[iCount];

                                    if (shape != null && shape.OLEFormat != null) {
                                        bool bRetVal = false;
                                        string strProgID;
                                        Guid autoConvert;
                                        strProgID = shape.OLEFormat.ProgID;
                                        bRetVal = GetFinalCLSID(ref strProgID, out autoConvert);

                                        // if we are successful with the conversion of the CLSID we now need to query
                                        //  the application to see if it can actually do the work
                                        if (bRetVal == true) {
                                            bool bInsertable = false;
                                            bool bNotInsertable = false;

                                            bInsertable = IsCLSIDInsertable(ref autoConvert);
                                            bNotInsertable = IsCLSIDNotInsertable(ref autoConvert);

                                            //Make sure that the server of interest is insertable and not-insertable
                                            if (bInsertable && bNotInsertable) {
                                                bool bServerExists = false;
                                                string strPathToExe;
                                                bServerExists = DoesServerExist(out strPathToExe, ref autoConvert);

                                                //if the server exists then see if MathML can be retrieved for the shape
                                                if (bServerExists) {
                                                    bool bMathML = false;
                                                    string strVerb;
                                                    int indexForVerb = -100;

                                                    strVerb = "RunForConversion";

                                                    bMathML = DoesServerSupportMathML(ref autoConvert, ref strVerb, out indexForVerb);
                                                    if (bMathML) {
                                                        storeMathMLEquation(ref shape, indexForVerb, listmathML);
                                                    }
                                                }
                                            } else {
                                                if (bInsertable != bNotInsertable) {
                                                    showMsg = 1;
                                                }
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                            //Increment the LCV and the number of shapes that iterated over.
                            iCount++;
                            iNumShapesViewed++;
                        }
                    }
                    rng = rng.NextStoryRange;
                }
                mathMLEquationsTable.Add(storyName, listmathML);
            }
            if (showMsg == 1) {
                string message =
                    "In order to convert MathType or Microsoft Equation Editor equations to DAISY,MathType 6.5 or later must be installed. See www.dessci.com/saveasdaisy for further information.Currently all the equations will be converted as Images";
                eventsHandler.OnStop(message, "Warning");
            }
            return iNumShapesViewed;
        }


        static private bool GetFinalCLSID(ref string ProgID, out Guid finalCLSID) {
            bool bRetVal = false;
            Guid oGuid;
            int iCOMRetVal = 0;

            iCOMRetVal = CLSIDFromProgID(ProgID, out oGuid);
            
            if (iCOMRetVal == 0) { // S_OK => Prog is associated to a Class
                FindFinalCLSID(ref oGuid, out finalCLSID);
                bRetVal = true;
            } else {
                finalCLSID = oGuid;
            }

            return bRetVal;
        }

        /// <summary>
        /// If a class belong to an "auto-conversion" chain of classes in the registry,
        /// get the final class of the chain.
        /// </summary>
        /// <param name="oldCLSID">original class ID</param>
        /// <param name="newCLSID">CLSID of the class that is at the end of the auto-conversion chain</param>
        static private void FindFinalCLSID(ref Guid oldCLSID, out Guid newCLSID) {
            int iCOMRetVal = 0;
            // Check if the class pointed by oldCLSID is set to be automatically to an other CLSID
            iCOMRetVal = OleGetAutoConvert(ref oldCLSID, out newCLSID);
            if (iCOMRetVal == 0) { // S_OK : no error during the call
                // Check if the new CLSID is not the old one (meaning we found the final CLSID)
                bool bGuidTheSame = false;
                try {
                    bGuidTheSame = IsEqualGUID(ref oldCLSID, ref newCLSID);
                } catch (COMException eCOM) {
                    MessageBox.Show(eCOM.Message);
                } catch (Exception e) {
                    MessageBox.Show(e.Message);
                }
                // Not the same ? keep going with the search
                if (bGuidTheSame == false) {
                    oldCLSID = newCLSID;
                    FindFinalCLSID(ref oldCLSID, out newCLSID);
                }
            } else {
                //There was some error in the auto conversion.
                // See if this guid will do the conversion.
                newCLSID = oldCLSID;
            }
        }

        static private bool IsCLSIDInsertable(ref Guid oGuid) {
            bool bInsertable = false;
            //Check for the existance of the insertable key
            RegistryKey regkey;/* new Microsoft.Win32 Registry Key */
            string strRegLocation;

            strRegLocation = @"Software\Classes\CLSID\" + @"{" + oGuid.ToString() + @"}" + @"\" + @"Insertable";
            regkey = Registry.LocalMachine.OpenSubKey(strRegLocation);

            if (regkey != null)
                bInsertable = true;

            return bInsertable;
        }

        static private bool IsCLSIDNotInsertable(ref Guid oGuid) {
            bool bNotInsertable = false;
            //Check for the existance of the insertable key
            RegistryKey regkey;/* new Microsoft.Win32 Registry Key */
            string strRegLocation;
            strRegLocation = @"Software\Classes\CLSID\" + @"{" + oGuid.ToString() + @"}" + @"\" + @"NotInsertable";

            regkey = Registry.LocalMachine.OpenSubKey(strRegLocation);

            //The not-insertable key is not present.
            if (regkey == null)
                bNotInsertable = true;

            return bNotInsertable;
        }

        static private bool DoesServerExist(out string strPathToExe, ref Guid oGuid) {
            bool bServerExists = false;
            //Check for the existance of the insertable key
            RegistryKey regkey;/* new Microsoft.Win32 Registry Key */
            string strRegLocation;
            strRegLocation = @"Software\Classes\CLSID\" + @"{" + oGuid.ToString() + @"}" + @"\" + @"LocalServer32";
            regkey = Registry.LocalMachine.OpenSubKey(strRegLocation);

            if (regkey != null) {
                string[] valnames = regkey.GetValueNames();
                strPathToExe = "";
                try {
                    strPathToExe = (string)regkey.GetValue(valnames[0]);
                } catch (Exception e) {
                }

                if (strPathToExe.Length > 0) {
                    //Now check if this is a good path
                    if (File.Exists(strPathToExe))
                        bServerExists = true;
                }

            } else {

                strPathToExe = null;

            }

            return bServerExists;
        }

        static private bool DoesServerSupportMathML(ref Guid oGuid, ref string strVerb, out int indexForVerb) {
            bool bIsMathMLSupported = false;
            //Check for the existance of the insertable key
            RegistryKey regkey;
            string strRegLocation;
            strRegLocation = @"Software\Classes\CLSID\" + "{" + oGuid.ToString() + "}" + @"\DataFormats\GetSet";
            regkey = Registry.LocalMachine.OpenSubKey(strRegLocation);

            if (regkey != null) {
                string[] valnames = regkey.GetSubKeyNames();
                int x = 0;
                while (x < regkey.SubKeyCount) {
                    RegistryKey subKey;
                    if (regkey.SubKeyCount > 0) {
                        subKey = regkey.OpenSubKey(valnames[x]);
                        if (subKey != null) {
                            string[] dataFormats = subKey.GetValueNames();
                            int y = 0;
                            while (y < subKey.ValueCount) {
                                string strValue = (string)subKey.GetValue(dataFormats[y]);

                                //This will accept both MathML and MathML Presentation.
                                if (strValue.Contains("MathML")) {
                                    bIsMathMLSupported = true;
                                    break;
                                }
                                y++;
                            }
                        }
                    }

                    if (bIsMathMLSupported)
                        break;
                    x++;
                }
            }

            //Now lets check to see if the appropriate verb is supported
            if (bIsMathMLSupported) {
                //The return value for a verb not found will be 1000
                //
                indexForVerb = GetVerbIndex(strVerb, ref oGuid);

                if (indexForVerb == 1000) {
                    bIsMathMLSupported = false;
                }
            } else {
                //We do not have an appropriate verb to start the server
                indexForVerb = -100;  //There is a predefined range for 
            }

            return bIsMathMLSupported;
        }

        static private int GetVerbIndex(string strVerbToFind, ref Guid oGuid) {
            int indexForVerb = 1000;
            //Check for the existance of the insertable key
            RegistryKey regkey;
            string strRegLocation;
            strRegLocation = @"Software\Classes\CLSID\" + "{" + oGuid.ToString() + "}" + @"\Verb";
            regkey = Registry.LocalMachine.OpenSubKey(strRegLocation);

            if (regkey != null) {
                //Lets make sure that we have some values before preceeding.
                if (regkey.SubKeyCount > 0) {
                    int x = 0;
                    int iCount = 0;

                    string[] valnames = regkey.GetSubKeyNames();

                    while (x < regkey.SubKeyCount) {
                        RegistryKey subKey;
                        if (regkey.SubKeyCount > 0) {
                            subKey = regkey.OpenSubKey(valnames[x]);
                            if (subKey != null) {
                                int y = 0;
                                string[] verbs = subKey.GetValueNames();
                                iCount = subKey.ValueCount;
                                string verb;

                                //Search all of the verbs for requested string.
                                while (y < iCount) {
                                    verb = (string)subKey.GetValue(verbs[y]);
                                    if (verb.Contains(strVerbToFind) == true) {
                                        string numVerb;
                                        numVerb = valnames[x].ToString();
                                        indexForVerb = int.Parse(numVerb);
                                        break;
                                    }
                                    y++;
                                }
                            }
                        }

                        //If the verb is not 1000 then break out of the verb
                        if (indexForVerb != 1000)
                            break;

                        x++;
                    }
                }
            }


            return indexForVerb;
        }

        static private void storeMathMLEquation(
            ref Microsoft.Office.Interop.Word.InlineShape shape,
            int indexForVerb,
            ArrayList listmathML
        ) {
            IConnectDataObject mDataObject;
            if (shape != null) {
                object dataObject = null;
                object objVerb;

                objVerb = indexForVerb;

                //Start MathType, and get the dataobject that is connected to the server.    
                shape.OLEFormat.DoVerb(ref objVerb);

                try {
                    dataObject = shape.OLEFormat.Object;
                } catch (Exception e) {
                    //we have an issue with trying to get the verb,
                    //  There will be a attempt at another way to start the application.
                    MessageBox.Show(e.Message);
                }

                IOleObject oleObject = null;

                //This is a C# version of a QueryInterface
                if (dataObject != null) {
                    mDataObject = dataObject as IConnectDataObject;
                    oleObject = dataObject as IOleObject;
                } else {
                    //There was an issue with the addin trying to start with the verb we
                    // knew.  A backup is to call the with the primary verb and start the 
                    //  application normally.
                    objVerb = MSword.WdOLEVerb.wdOLEVerbPrimary;
                    shape.OLEFormat.DoVerb(ref objVerb);

                    dataObject = shape.OLEFormat.Object;
                    mDataObject = dataObject as IConnectDataObject;
                    oleObject = dataObject as IOleObject;
                }
                //Create instances of FORMATETC and STGMEDIUM for use with IDataObject
                ConnectFORMATETC oFormatEtc = new ConnectFORMATETC();
                ConnectSTGMEDIUM oStgMedium = new ConnectSTGMEDIUM();
                DataFormats.Format oFormat;



                //Find within the clipboard system the registered clipboard format for MathML
                oFormat = DataFormats.GetFormat("MathML");

                if (mDataObject != null) {
                    int iRetVal = 0;

                    //Initialize a FORMATETC structure to get the requested data
                    oFormatEtc.cfFormat = (Int16)oFormat.Id;
                    oFormatEtc.dwAspect = System.Runtime.InteropServices.ComTypes.DVASPECT.DVASPECT_CONTENT;
                    oFormatEtc.lindex = -1;
                    oFormatEtc.ptd = (IntPtr)0;
                    oFormatEtc.tymed = TYMED.TYMED_HGLOBAL;

                    iRetVal = mDataObject.QueryGetData(ref oFormatEtc);
                    //iRetVal will be zero if the MathML type is contained within the server.
                    if (iRetVal == 0) {
                        oStgMedium.tymed = TYMED.TYMED_NULL;
                    }

                    try {
                        mDataObject.GetData(ref oFormatEtc, out oStgMedium);
                    } catch (System.Runtime.InteropServices.COMException e) {
                        System.Windows.Forms.MessageBox.Show(e.ToString());
                        throw;
                    }

                    // Because we explicitly requested a MathML, we know that it is TYMED_HGLOBAL
                    // lets deal with the memory here.
                    if (oStgMedium.tymed == TYMED.TYMED_HGLOBAL &&
                        oStgMedium.unionmember != null) {
                        WriteOutMathMLFromStgMedium(ref oStgMedium, listmathML);

                        if (oleObject != null) {
                            uint close = (uint)OLECLOSE.OLECLOSE_NOSAVE;
                            // uint close = (uint)Microsoft.VisualStudio.OLE.Interop.OLECLOSE.OLECLOSE_NOSAVE;
                            oleObject.Close(close);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Store the mathML formated equation in a provided mathMl string array
        /// </summary>
        /// <param name="oStgMedium">memory handler of the equation</param>
        /// <param name="listmathML">String array to store the mathml in</param>
        static private void WriteOutMathMLFromStgMedium(
            ref ConnectSTGMEDIUM oStgMedium,
            ArrayList listmathML
        ) {
            IntPtr ptr;
            byte[] rawArray = null;


            //Verify that our data contained within the STGMEDIUM is non-null
            if (oStgMedium.unionmember != null) {
                //Get the pointer to the data that is contained
                //  within the STGMEDIUM
                ptr = oStgMedium.unionmember;

                //The pointer now becomes a Handle reference.
                HandleRef handleRef = new HandleRef(null, ptr);

                try {
                    //Lock in the handle to get the pointer to the data
                    IntPtr ptr1 = GlobalLock(handleRef);

                    //Get the size of the memory block
                    int length = GlobalSize(handleRef);

                    //New an array of bytes and Marshal the data across.
                    rawArray = new byte[length];
                    Marshal.Copy(ptr1, rawArray, 0, length);

                    // I will now display the text.  Create a string from the rawArray
                    string str = Encoding.ASCII.GetString(rawArray);
                    str = str.Substring(str.IndexOf("<mml:math"), str.IndexOf("</mml:math>") - str.IndexOf("<mml:math"));
                    str = str + "</mml:math>";
                    str = str.Replace("xmlns:mml='http://www.w3.org/1998/Math/MathML'", "");

                    listmathML.Add(str);
                } catch (Exception exp) {
                    System.Diagnostics.Debug.WriteLine("MathMLimport from MathType threw an exception: " + Environment.NewLine + exp.ToString());
                } finally {
                    //This gets called regardless within a try catch.
                    //  It is a good place to clean up like this.
                    GlobalUnlock(handleRef);
                }
            }
        }
        #endregion // MathML


        #region SubDocuments

        

        /// <summary>
        /// Prepare conversion of subdocuments
        /// </summary>
        /// <param name="eventsHandler"></param>
        /// <param name="preprocessingResult"></param>
        /// <param name="master">Opened master document</param>
        /// <returns>true if the preprocessinf completed successfully</returns>
        public static bool preprocessSubDocuments(
            IConversionEventsHandler eventsHandler,
            PreprocessingData preprocessingResult,
            Document master
        ) {
            SubdocumentsList subdocuments = SubdocumentsManager.FindSubdocuments(
                preprocessingResult.conversionParameters.TempInputFile, 
                preprocessingResult.conversionParameters.InputFile);
            //MessageBox.Show("Check for errors when retrieving subdocuments pathes");
            if (subdocuments.Errors.Count > 0) {
                StringBuilder errorMessage = new StringBuilder();
                errorMessage.Append("Errors were encoutered while retrieving sub documents:");
                foreach (string error in subdocuments.Errors) {
                    errorMessage.Append("\r\n- " + error);
                }
                eventsHandler.OnError(errorMessage.ToString());
                return false;
            }

            //MessageBox.Show("Get not translated docs");
            //notTranslatedDoc = subdocuments.GetNotTraslatedSubdocumentsNames();

            //MessageBox.Show("Doc sublist");
            ArrayList subList = new ArrayList();
            subList.Add(preprocessingResult.conversionParameters.TempInputFile + "|Master");
            foreach (string subdoc in subdocuments.GetSubdocumentsNameWithRelationship()) {
                subList.Add(subdoc);
            }

            int subCount = subdocuments.SubdocumentsCount + 1;
            //Checking whether any original or Subdocumets is already Open or not
            foreach (string docPathAndType in subList) {
                string[] splitted = docPathAndType.Split('|');
                if (PackageUtilities.documentIsOpen(splitted[0])) {
                    eventsHandler.OnError("Some Sub documents are in open state.\r\nPlease close all the Sub documents before Translation.");
                    return false;
                }
            }

            //Checking whether Sub documents are Simple documents or a Master document
            string resultSub = SubdocumentsManager.CheckingSubDocs(subdocuments.GetSubdocumentsNameWithRelationship());
            if (resultSub != "simple") {
                eventsHandler.OnError("Some of the added documents are MasterSub documents.Please add simple documents.");
                return false;
            }

            List<string> warnings = new List<string>();
            MSword.Application WordInstance = master.Application;
            for (int i = 0; i < subList.Count; i++) {
                string[] splitName = subList[i].ToString().Split('|');
                object addToRecentFiles = false;
                object readOnly = false;
                object isVisible = false;
                object missing = Type.Missing;
                object saveChanges = MSword.WdSaveOptions.wdDoNotSaveChanges;
                object originalFormat = MSword.WdOriginalFormat.wdOriginalDocumentFormat;
                object newName = splitName[0].ToString();

                MSword.Document subDoc = WordInstance.Documents.Open(ref newName, ref missing, ref readOnly, ref addToRecentFiles, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);
                try {
                    WordPreprocessing.exportShapes(
                        subDoc,
                        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + "SaveAsDAISY");
                } catch (Exception e) {
                    warnings.Add(e.Message);
                }
                try {
                    WordPreprocessing.exportInlineShapes(
                       subDoc,
                       Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + "SaveAsDAISY");
                } catch (Exception e) {
                    warnings.Add(e.Message);
                }
                Hashtable multipleOwnMathMl = new Hashtable();
                WordPreprocessing.parseEquations(
                    eventsHandler,
                    subDoc,
                    multipleOwnMathMl
                );
                preprocessingResult.conversionParameters.ListMathMl.Add("Doc" + i, multipleOwnMathMl);

                subDoc.Close(saveChanges, originalFormat);
            }
            if (warnings.Count > 0) {
                string warningMessage = "Errors occured while preprocessing subdocuments:\r\n";
                foreach (string warning in warnings) {
                    warningMessage += "\r\n" + warning;
                }
                eventsHandler.OnError(warningMessage);
            }

            //this.applicationObject.ActiveDocument.Save();
            return true;
        }

        #endregion

        


    }
}
