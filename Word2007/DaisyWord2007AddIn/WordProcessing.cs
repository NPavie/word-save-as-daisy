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

using Daisy.DaisyConverter.DaisyConverterLib;
using System.Drawing.Imaging;

namespace DaisyWord2007AddIn {

    /// <summary>
    /// MS Word processing functions meant to be used in external programs
    /// </summary>
    static class WordProcessing {

        #region Shapes
        /// <summary>
        /// Save the non-textual shapes of a list of document as png images
        /// (shapes that do not contain "Text Box" in its name)
        /// 
        /// </summary>
        static public void saveShapes(
            MSword.Application WordInstance,
            ArrayList documentsPathes,
            string outputPath
        ) {
            object addToRecentFiles = false;
            object readOnly = false;
            object isVisible = false;
            object missing = Type.Missing;
            object saveChanges = MSword.WdSaveOptions.wdDoNotSaveChanges;
            object originalFormat = MSword.WdOriginalFormat.wdOriginalDocumentFormat;
            System.Diagnostics.Process objProcess = System.Diagnostics.Process.GetCurrentProcess();
            List<string> warnings = new List<string>();
            for (int i = 0; i < documentsPathes.Count; i++) {
                object newName = documentsPathes[i];
                String fileName = newName.ToString().Replace(" ", "_");
                MSword.Document newDoc = WordInstance.Documents.Open(ref newName, ref missing, ref readOnly, ref addToRecentFiles, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);
                foreach (MSword.Shape item in newDoc.Shapes) {
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

                newDoc.Close(ref saveChanges, ref originalFormat, ref missing);
            }
            if (warnings.Count > 0) {
                string warningMessage = "Some shapes could not be exported from the documents:";
                foreach (string warning in warnings) {
                    warningMessage += "\r\n" + warning;
                }
                throw new Exception(warningMessage);
            }
        }

        /// <summary>
        /// Save the inline shapes for a list a documents
        /// (Not e : not of type Embedded OLE or Pictures, so it is probably targeting inlined vectorial drawings)
        /// </summary>
        static public void saveInlineShapes(
            MSword.Application WordInstance,
            ArrayList documentsPathes,
            string outputPath
        ) {
            System.Diagnostics.Process objProcess = System.Diagnostics.Process.GetCurrentProcess();
            object addToRecentFiles = false;
            object readOnly = false;
            object isVisible = false;
            object missing = Type.Missing;
            object saveChanges = MSword.WdSaveOptions.wdDoNotSaveChanges;
            object originalFormat = MSword.WdOriginalFormat.wdOriginalDocumentFormat;
            List<string> warnings = new List<string>();
            for (int i = 0; i < documentsPathes.Count; i++) {
                object newName = documentsPathes[i].ToString();
                MSword.Document newDoc = WordInstance.Documents.Open(ref newName, ref missing, ref readOnly, ref addToRecentFiles, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);
                MSword.Range rng;
                String fileName = newName.ToString().Replace(" ", "_");
                foreach (MSword.Range tmprng in newDoc.StoryRanges) {
                    rng = tmprng;
                    while (rng != null) {
                        foreach (MSword.InlineShape item in rng.InlineShapes) {
                            if ((item.Type.ToString() != "wdInlineShapeEmbeddedOLEObject") && ((item.Type.ToString() != "wdInlineShapePicture"))) {
                                string str = "Shapes_" + GenerateId().ToString();
                                string pathShape = outputPath + "\\" + Path.GetFileNameWithoutExtension(fileName) + "-" + str + ".png";

                                object range = item.Range;

                                item.Range.Bookmarks.Add(str, ref range);
                                item.Range.CopyAsPicture();

                                try {
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

                newDoc.Save();
                newDoc.Close(ref saveChanges, ref originalFormat, ref missing);
            }

            if (warnings.Count > 0) {
                string warningMessage = "Some images could not be exported from the documents:";
                foreach (string warning in warnings) {
                    warningMessage += "\r\n" + warning;
                }
                throw new Exception(warningMessage);
            }
        }

        #endregion

        #region MathML
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

        /// <summary>
        /// Search the equations in the stories of a document.
        /// </summary>
        /// <param name="eventsHandler"></param>
        /// <param name="WordInstance">Instance of the word application</param>
        /// <param name="documentPath">Path of the document to parse</param>
        /// <param name="mathMLEquationsTable">Table to store the equations list per document story</param>
        /// <returns>The number of equations found</returns>
        static public int parseEquations(
            IPluginEventsHandler eventsHandler,
            MSword.Application WordInstance,
            string documentPath,
            Hashtable mathMLEquationsTable
        ) {
            Int16 showMsg = 0;
            MSword.Range rng;
            String storyName = "";
            int iNumShapesViewed = 0;

            object addToRecentFiles = false;
            object readOnly = false;
            object isVisible = false;
            object missing = Type.Missing;
            object saveChanges = MSword.WdSaveOptions.wdDoNotSaveChanges;
            object originalFormat = MSword.WdOriginalFormat.wdOriginalDocumentFormat;
            object newName = documentPath;
            String fileName = newName.ToString().Replace(" ", "_");
            MSword.Document doc = WordInstance.Documents.Open(ref newName, ref missing, ref readOnly, ref addToRecentFiles, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);
            //Make sure that we have shapes to iterate over

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
                                        bRetVal = FindAutoConvert(ref strProgID, out autoConvert);

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
            //System.Windows.Forms.MessageBox.Show("In order to convert MathType or Microsoft Equation Editor equations to DAISY,MathType 6.5 or later must be installed. See www.dessci.com/saveasdaisy for further information.Currently all the equations will be converted as Images", "Warning", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
            doc.Close();
            return iNumShapesViewed;
        }


        static private bool FindAutoConvert(ref string ProgID, out Guid autoConvert) {
            bool bRetVal = false;
            Guid oGuid;
            int iCOMRetVal = 0;

            iCOMRetVal = CLSIDFromProgID(ProgID, out oGuid);

            if (iCOMRetVal == 0) {
                RecurseAutoConvert(ref oGuid, out autoConvert);
                bRetVal = true;
            } else {
                autoConvert = oGuid;
            }

            return bRetVal;
        }

        static private void RecurseAutoConvert(ref Guid oGuid, out Guid autoConvert) {
            int iCOMRetVal = 0;

            iCOMRetVal = OleGetAutoConvert(ref oGuid, out autoConvert);
            if (iCOMRetVal == 0) {
                //If we have no error and the the CLSIDs are the same, then make sure that
                // 
                bool bGuidTheSame = false;
                try {
                    bGuidTheSame = IsEqualGUID(ref oGuid, ref autoConvert);
                } catch (COMException eCOM) {
                    MessageBox.Show(eCOM.Message);
                } catch (Exception e) {
                    MessageBox.Show(e.Message);
                }

                if (bGuidTheSame == false) {
                    oGuid = autoConvert;
                    RecurseAutoConvert(ref oGuid, out autoConvert);
                }
            } else {
                //There was some error in the auto conversion.
                // See if this guid will do the conversion.
                autoConvert = oGuid;
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

        /// <summary>
        /// Function to generate Random ID
        /// </summary>
        /// <returns></returns>
        static public long GenerateId() {
            byte[] buffer = Guid.NewGuid().ToByteArray();
            return BitConverter.ToInt64(buffer, 0);
        }
    }
}
