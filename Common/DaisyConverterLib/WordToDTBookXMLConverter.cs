/* 
 * Copyright (c) 2006, Clever Age
 * All rights reserved.
 * 
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of Clever Age nor the names of its contributors 
 *       may be used to endorse or promote products derived from this software
 *       without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY THE REGENTS AND CONTRIBUTORS ``AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL THE REGENTS AND CONTRIBUTORS BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

using System;
using System.Diagnostics;
using System.IO;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Xsl;
using System.Reflection;
using System.Collections;
using System.Xml.Schema;
using System.Windows.Forms;
using System.IO.Packaging;
using System.Text;

namespace Daisy.SaveAsDAISY.DaisyConverterLib
{
    /// <summary>
    /// Core conversion class that convert a word file (on disk) to XML
    /// </summary>
    public class WordToDTBookXMLConverter
    {
        private const string SOURCE_XML = "source.xml";

        private bool isDirectTransform = true;
        private ArrayList skipedPostProcessors = null;
        private string externalResource = null;
        private Assembly resourcesAssembly;
        private Hashtable compiledProcessors;
        private static bool isValid;
        private static string error;

        ArrayList MathEntities8879, MathEntities9573, MathMLEntities;

        private ArrayList fidilityLoss = new ArrayList();

        String errorText = "";

        ChainResourceManager resourceManager;

        #region Fields
        public bool DirectTransform {
            set { this.isDirectTransform = value; }
            get { return this.isDirectTransform; }
        }

        public ArrayList SkipedPostProcessors {
            set { this.skipedPostProcessors = value; }
        }

        public string ExternalResources {
            set { this.externalResource = value; }
            get { return this.externalResource; }
        }

        public ArrayList FidilityLoss {
            get {
                return fidilityLoss;
            }
        }

        /// <summary>
        /// Pull the chain of post processors for the direct conversion
        /// </summary>
        protected virtual string[] DirectPostProcessorsChain {
            get { return null; }
        }

        /// <summary>
        /// Pull the chain of post processors for the reverse conversion
        /// </summary>
        protected virtual string[] ReversePostProcessorsChain {
            get { return null; }
        }

        /// <summary>
        /// Pull an XmlUrlResolver for embedded resources
        /// </summary>
        private XmlUrlResolver ResourceResolver {
            get {
                if (this.ExternalResources == null) {
                    return new EmbeddedResourceResolver(this.resourcesAssembly,
                      this.GetType().Namespace, this.DirectTransform);
                } else {
                    return new SharedXmlUrlResolver(this.DirectTransform);
                }
            }
        }

        /// <summary>
        /// Pull the input xml document to the xsl transformation
        /// </summary>
        private XmlReader Source {
            get {
                XmlReaderSettings xrs = new XmlReaderSettings();
                // do not look for DTD
                xrs.DtdProcessing = DtdProcessing.Prohibit;
                if (this.ExternalResources == null) {
                    xrs.XmlResolver = this.ResourceResolver;
                    return XmlReader.Create(SOURCE_XML, xrs);
                } else {
                    return XmlReader.Create(this.ExternalResources + "/" + SOURCE_XML, xrs);
                }
            }
        }


        /// <summary>
        /// Pull the xslt settings
        /// </summary>
        private XsltSettings XsltProcSettings {
            get {
                // Enable xslt 'document()' function
                return new XsltSettings(true, false);
            }
        }
        #endregion

        public WordToDTBookXMLConverter(
            System.Resources.ResourceManager customResourceManager = null
        ) {
            this.resourcesAssembly = Assembly.GetExecutingAssembly();
            this.skipedPostProcessors = new ArrayList();
            this.compiledProcessors = new Hashtable();

            this.resourceManager = new ChainResourceManager();
            
            // Add a default resource managers (for common labels)
            this.resourceManager.Add(
                new System.Resources.ResourceManager("Daisy.SaveAsDAISY.DaisyConverterLib.resources.Labels",
                Assembly.GetExecutingAssembly()));

            // additionnal resource manager
            if(customResourceManager != null) {
                this.resourceManager.Add(customResourceManager);
            }

            // Load math and mathml entities
            MathEntities8879 = new ArrayList();
            MathEntities9573 = new ArrayList();
            MathMLEntities = new ArrayList();

            MathEntities8879.Add("isobox.ent");
            MathEntities8879.Add("isocyr1.ent");
            MathEntities8879.Add("isocyr2.ent");
            MathEntities8879.Add("isodia.ent");
            MathEntities8879.Add("isolat1.ent");
            MathEntities8879.Add("isolat2.ent");
            MathEntities8879.Add("isonum.ent");

            MathEntities8879.Add("isopub.ent");

            MathMLEntities.Add("mmlalias.ent");
            MathMLEntities.Add("mmlextra.ent");

            MathEntities9573.Add("isoamsa.ent");
            MathEntities9573.Add("isoamsb.ent");
            MathEntities9573.Add("isoamsc.ent");
            MathEntities9573.Add("isoamsn.ent");
            MathEntities9573.Add("isoamso.ent");
            MathEntities9573.Add("isoamsr.ent");
            MathEntities9573.Add("isogrk3.ent");
            MathEntities9573.Add("isomfrk.ent");
            MathEntities9573.Add("isomopf.ent");
            MathEntities9573.Add("isomscr.ent");
            MathEntities9573.Add("isotech.ent");

        }

        private XslCompiledTransform Load(bool computeSize)
        {
            try
            {
                string xslLocation = "oox2Daisy.xsl";
                XPathDocument xslDoc = null;
                XmlUrlResolver resolver = this.ResourceResolver;

                if (this.ExternalResources == null)
                {
                    if (computeSize)
                    {
                        xslLocation = "oox2Daisy.xsl";
                    }
                    EmbeddedResourceResolver emr = (EmbeddedResourceResolver)resolver;
                    emr.IsDirectTransform = this.DirectTransform;
                    xslDoc = new XPathDocument(emr.GetInnerStream(xslLocation));
                }
                else
                {
                    xslDoc = new XPathDocument(this.ExternalResources + "/" + xslLocation);
                }

                if (!this.compiledProcessors.Contains(xslLocation))
                {

#if DEBUG
                    XslCompiledTransform xslt = new XslCompiledTransform(true);
#else
                    XslCompiledTransform xslt = new XslCompiledTransform();
#endif

                    // compile the stylesheet. 
                    // Input stylesheet, xslt settings and uri resolver are retrieve from the implementation class.
                    xslt.Load(xslDoc, this.XsltProcSettings, this.ResourceResolver);
                    this.compiledProcessors.Add(xslLocation, xslt);
                }
                return (XslCompiledTransform)this.compiledProcessors[xslLocation];
            }
            catch (Exception e)
            {
                string s;
                s = e.Message;
                return null;
            }

        }


        public void ComputeSize(string inputFile, Hashtable table)
        {
            Transform(inputFile, null, table, null, true, "");
        }

        /// <summary>
        /// Important function to be documented
        /// </summary>
        /// <param name="inputFile"></param>
        /// <param name="outputFile"></param>
        /// <param name="table"></param>
        /// <param name="listMathMl"></param>
        /// <param name="modeValue"></param>
        /// <param name="output_Pipeline"></param>
        public void Transform(string inputFile, string outputFile, Hashtable table, Hashtable listMathMl, bool modeValue, string output_Pipeline)
        {
            fidilityLoss = new ArrayList();
            string tempInputFile = Path.GetTempFileName();
            string tempOutputFile = outputFile == null ? null : Path.GetTempFileName();
            try
            {
                File.Copy(inputFile, tempInputFile, true);
                File.SetAttributes(tempInputFile, FileAttributes.Normal);
                _Transform(inputFile, tempInputFile, tempOutputFile, outputFile, listMathMl, table, output_Pipeline);

                if (outputFile != null)
                {
                    if (File.Exists(outputFile))
                    {
                        File.Delete(outputFile);
                    }
                    File.Move(tempOutputFile, outputFile);

                    CopyDTDToDestinationfolder(outputFile);
                    CopyCSSToDestinationfolder(outputFile);
                    CopyMATHToDestinationfolder(outputFile);

                    if (modeValue)
                    {
                        XmlValidation(outputFile);

                    }

                    Int16 value = (Int16)outputFile.LastIndexOf("\\");
                    String tempStr = outputFile.Substring(0, value);
                    DeleteDTD(tempStr + "\\" + "dtbook-2005-3.dtd", outputFile, modeValue);
                    DeleteMath(tempStr, modeValue);
                }
            }
            finally
            {
                if (File.Exists(tempInputFile))
                {
                    try
                    {
                        File.Delete(tempInputFile);
                    }
                    catch (IOException)
                    {
                        Debug.Write("could not delete temporary input file");
                    }
                }
            }
        }

        /// <summary>
        /// Convert a single word document to DTBook XML using xslts
        /// </summary>
        /// <param name="inputDocumentPath">Path of the Word document</param>
        /// <param name="outputPath">Output path where the resulting xml, dtds </param>
        /// <param name="conversion">Current conversion parameters</param>
        /// <param name="validate">If true, the resulting xml is also validated against the dtds</param>
        /// <param name="mathMLSubSetKey">If defined, a subset of mathml entries are used for the conversion</param>
        public void convert(string inputDocumentPath, string outputPath, ConversionParameters conversion, bool validate, string mathMLSubSetKey = "") {
            fidilityLoss = new ArrayList();

            // temporary files
            string tempInputPath = Path.GetTempFileName();
            string tempOutputPath = outputPath == null ? null : Path.GetTempFileName();

            try {
                File.Copy(inputDocumentPath, tempInputPath, true);
                File.SetAttributes(tempInputPath, FileAttributes.Normal);

                XmlReader source = null;
                XmlWriter writer = null;
                ZipResolver zipResolver = null;

                try {
                    XslCompiledTransform xslt = this.Load(tempOutputPath == null);
                    zipResolver = new ZipResolver(conversion.TempInputFile);

                    XsltArgumentList parameters = new XsltArgumentList();
                    parameters.XsltMessageEncountered += new XsltMessageEncounteredEventHandler(onXSLTMessageEvent);
                    parameters.XsltMessageEncountered += new XsltMessageEncounteredEventHandler(onXSLTProgressMessageEvent);

                    string conversionOutput = tempOutputPath == null ?
                        tempInputPath.Substring(0, tempInputPath.LastIndexOf("\\")) :
                        (outputPath.ToLower().EndsWith(".xml") ?
                            outputPath.Substring(0, outputPath.LastIndexOf("\\")) :
                            outputPath);

                    DaisyClass val = new DaisyClass(
                                inputDocumentPath,
                                tempInputPath,
                                conversionOutput,
                                mathMLSubSetKey != "" ? (Hashtable) conversion.ListMathMl[mathMLSubSetKey] : conversion.ListMathMl,
                                zipResolver.Archive,
                                conversion.PipelineOutput);
                    parameters.AddExtensionObject("urn:Daisy", val);


                    parameters.AddParam("outputFile", "", tempOutputPath);

                    foreach (DictionaryEntry myEntry in conversion.ConversionParametersHash) {
                        parameters.AddParam(myEntry.Key.ToString(), "", myEntry.Value.ToString());
                    }

                    if (conversion.TempOutputFile != null) {
                        XmlWriter finalWriter;
#if DEBUG
                        StreamWriter streamWriter = new StreamWriter(
                            conversion.TempOutputFile,
                            true, System.Text.Encoding.UTF8
                        ) { AutoFlush = true };
                        finalWriter = new XmlTextWriter(streamWriter);
                        Debug.WriteLine("OUTPUT FILE : '" + conversion.TempOutputFile + "'");
#else
                        finalWriter = new XmlTextWriter(conversion.TempOutputFile, System.Text.Encoding.UTF8);
#endif
                        writer = GetWriter(finalWriter);
                    } else {

                        writer = new XmlTextWriter(new StringWriter());
                    }

                    source = this.Source;
                    // Apply the transformation

                    xslt.Transform(source, parameters, writer, zipResolver);
                } finally {
                    if (writer != null)
                        writer.Close();
                    if (source != null)
                        source.Close();
                    if (zipResolver != null)
                        zipResolver.Dispose();
                }


                if (conversion.OutputFile != null) {
                    if (File.Exists(conversion.OutputFile)) {
                        File.Delete(conversion.OutputFile);
                    }
                    File.Move(conversion.TempOutputFile, conversion.OutputFile);
                    CopyCSSToDestinationfolder(conversion.OutputFile);
                    // TODO: The handling of DTD and math needs to be cleanedup
                    Int16 value = (Int16)conversion.OutputFile.LastIndexOf("\\");
                    String tempStr = conversion.OutputFile.Substring(0, value);
                    CopyDTDToDestinationfolder(conversion.OutputFile);
                    CopyMATHToDestinationfolder(conversion.OutputFile);
                    if (validate) {
                        XmlValidation(conversion.OutputFile);
                    }
                    // We need to change this method : it does not only delete the dtds files,
                    // it updates the file dtds
                    DeleteDTD(tempStr + "\\" + "dtbook-2005-3.dtd", conversion.OutputFile, validate);
                    DeleteMath(tempStr, validate);
                }
            } finally {
                if (File.Exists(conversion.TempInputFile)) {
                    try {
                        File.Delete(conversion.TempInputFile);
                    } catch (IOException) {
                        Debug.Write("could not delete temporary input file");
                    }
                }
            }
        }

        /// <summary>
        /// This function does the following action on the "filename" file : 
        /// - replace the system local dtbook-2005-3.dtd by the public one
        /// - add the correct namespace to the dtbook declaration
        ///
        /// Notes : 
        /// - if the value parameter is set to false, the function is only doing a read and rewrite the file on itself
        /// - If no closing mml:math tag is found, the code removes 
        ///   - chars 203 to 1120, 
        ///   - an empty line before the dtbook tag 
        ///   - every occurences of the mml namespace declaration
        /// </summary>
        /// <param name="fileDTD">dtd file to be removed from disk</param>
        /// <param name="fileName">path of the file to be updated</param>
        /// <param name="value">boolean flag - if true, tha actions are applied on the updated file</param>
        public void DeleteDTD(String fileDTD, String fileName, bool value)
        {

            /*temperory solution - needs to be changed*/
            StreamReader reader = new StreamReader(fileName);
            string data = reader.ReadToEnd();
            reader.Close();

            StreamWriter writer = new StreamWriter(fileName);
            if (value)
            {
                data = data.Replace("<!DOCTYPE dtbook SYSTEM 'dtbook-2005-3.dtd'", "<!DOCTYPE dtbook PUBLIC '-//NISO//DTD dtbook 2005-3//EN' 'http://www.daisy.org/z3986/2005/dtbook-2005-3.dtd'");
                data = data.Replace("<dtbook version=\"" + "2005-3\"", "<dtbook xmlns=\"http://www.daisy.org/z3986/2005/dtbook/\" version=\"2005-3\"");
                if (!data.Contains("</mml:math>"))
                {
                    data = data.Remove(203, 917);
                    data = data.Replace(Environment.NewLine + "<dtbook", "<dtbook");
                    data = data.Replace("xmlns:mml=\"http://www.w3.org/1998/Math/MathML\"", "");
                }
                data = data.Replace("<!ENTITY % mathML2 SYSTEM 'mathml2.dtd'>", "<!ENTITY % mathML2 PUBLIC \"-//W3C//DTD MathML 2.0//EN\" \"http://www.w3.org/Math/DTD/mathml2/mathml2.dtd\">");
            }
            writer.Write(data);
            writer.Close();

            if (value)
            {
                if (File.Exists(fileDTD))
                {
                    File.Delete(fileDTD);
                }
            }
        }
        /// <summary>
        /// Replace system DTDs (dtbook and mathml) by public ones in a xml document
        /// </summary>
        /// <param name="fileName">XML document path</param>
        public void switchToPublicDTDs(string fileName) {
            StreamReader reader = new StreamReader(fileName);
            string data = reader.ReadToEnd();
            reader.Close();

            StreamWriter writer = new StreamWriter(fileName);
            // Replace dtbook
            data = data.Replace("<!DOCTYPE dtbook SYSTEM 'dtbook-2005-3.dtd'", "<!DOCTYPE dtbook PUBLIC '-//NISO//DTD dtbook 2005-3//EN' 'http://www.daisy.org/z3986/2005/dtbook-2005-3.dtd'");
            data = data.Replace("<dtbook version=\"" + "2005-3\"", "<dtbook xmlns=\"http://www.daisy.org/z3986/2005/dtbook/\" version=\"2005-3\"");
            if (!data.Contains("</mml:math>")) { // remove mathml if no closing mml math tag is found
                data = data.Remove(203, 917); // FIXME : DANGEROUS hard coded deletion, need to identify what it deletes !
                data = data.Replace(Environment.NewLine + "<dtbook", "<dtbook");
                data = data.Replace("xmlns:mml=\"http://www.w3.org/1998/Math/MathML\"", "");
            }
            // Replace mathml
            data = data.Replace("<!ENTITY % mathML2 SYSTEM 'mathml2.dtd'>", "<!ENTITY % mathML2 PUBLIC \"-//W3C//DTD MathML 2.0//EN\" \"http://www.w3.org/Math/DTD/mathml2/mathml2.dtd\">");
            writer.Write(data);
            writer.Close();
        }

        public void DeleteMath(String fileName, bool value)
        {
            if (value)
            {
                DeleteFile(fileName + "\\mathml2.DTD");
                DeleteFile(fileName + "\\mathml2-qname-1.mod");
                Directory.Delete(fileName + "\\iso8879", true);
                Directory.Delete(fileName + "\\iso9573-13", true);
                Directory.Delete(fileName + "\\mathml", true);
            }
        }

        public void DeleteFile(String file)
        {
            if (File.Exists(file))
            {
                File.Delete(file);
            }
        }

        /// <summary>
        /// Imporant function to be documented !
        /// </summary>
        /// <param name="inputFile"></param>
        /// <param name="tempInputFile"></param>
        /// <param name="tempOutputFile"></param>
        /// <param name="outputFile"></param>
        /// <param name="listMathMl"></param>
        /// <param name="table"></param>
        /// <param name="output_Pipeline"></param>
        private void _Transform(String inputFile, string tempInputFile, string tempOutputFile, string outputFile, Hashtable listMathMl, Hashtable table, string output_Pipeline)
        {
            // this throws an exception in the the following cases:
            // - input file is not a valid file
            // - input file is an encrypted file

            XmlReader source = null;
            XmlWriter writer = null;
            ZipResolver zipResolver = null;

            try
            {
                XslCompiledTransform xslt = this.Load(tempOutputFile == null);
                zipResolver = new ZipResolver(tempInputFile);

                XsltArgumentList parameters = new XsltArgumentList();
                parameters.XsltMessageEncountered += new XsltMessageEncounteredEventHandler(onXSLTMessageEvent);
                parameters.XsltMessageEncountered += new XsltMessageEncounteredEventHandler(onXSLTProgressMessageEvent);

                if (tempOutputFile != null)
                {
                    if (outputFile.ToLower().EndsWith(".xml"))
                    {
                        int length = outputFile.LastIndexOf("\\");
                        string s1 = outputFile.Substring(0, length);


                        DaisyClass val = new DaisyClass(inputFile, tempInputFile, s1, listMathMl, zipResolver.Archive, output_Pipeline);
                        parameters.AddExtensionObject("urn:Daisy", val);


                        parameters.AddParam("outputFile", "", tempOutputFile);

                        foreach (DictionaryEntry myEntry in table)
                        {
                            parameters.AddParam(myEntry.Key.ToString(), "", myEntry.Value.ToString());
                        }


                        XmlWriter finalWriter;
#if DEBUG
						StreamWriter streamWriter = new StreamWriter(tempOutputFile, true, System.Text.Encoding.UTF8) {AutoFlush = true};
                    	finalWriter = new XmlTextWriter(streamWriter);
#else
                    	finalWriter = new XmlTextWriter(outputFile, System.Text.Encoding.UTF8);
#endif

#if DEBUG
                    	Debug.WriteLine("OUTPUT FILE : '" + tempOutputFile + "'");
#endif

                        writer = GetWriter(finalWriter);
                    }
                    else
                    {
                        DaisyClass val = new DaisyClass(inputFile, tempInputFile, outputFile, listMathMl, zipResolver.Archive, output_Pipeline);
                        parameters.AddExtensionObject("urn:Daisy", val);


                        parameters.AddParam("outputFile", "", tempOutputFile);

                        foreach (DictionaryEntry myEntry in table)
                        {
                            parameters.AddParam(myEntry.Key.ToString(), "", myEntry.Value.ToString());
                        }


                        XmlWriter finalWriter;
                        finalWriter = new XmlTextWriter(tempOutputFile, System.Text.Encoding.UTF8);

                        writer = GetWriter(finalWriter);
                    }

                }
                else
                {

                    int length = tempInputFile.LastIndexOf("\\");
                    string s1 = tempInputFile.Substring(0, length);

                    DaisyClass val = new DaisyClass(inputFile, tempInputFile, s1, listMathMl, zipResolver.Archive, output_Pipeline);
                    parameters.AddExtensionObject("urn:Daisy", val);

					if (table != null)
					{
						foreach (DictionaryEntry myEntry in table)
							parameters.AddParam(myEntry.Key.ToString(), "", myEntry.Value.ToString());
					}


                	writer = new XmlTextWriter(new StringWriter());
                }
                source = this.Source;
                // Apply the transformation

                xslt.Transform(source, parameters, writer, zipResolver);                
            }
            finally
            {
                if (writer != null)
                    writer.Close();
                if (source != null)
                    source.Close();
                if (zipResolver != null)
                    zipResolver.Dispose();
            }
        }


        

        private XmlWriter GetWriter(XmlWriter writer)
        {
            string[] postProcessors = this.DirectPostProcessorsChain;
            if (!this.isDirectTransform)
            {
                postProcessors = this.ReversePostProcessorsChain;
            }
            return InstanciatePostProcessors(postProcessors, writer);
        }


        private XmlWriter InstanciatePostProcessors(string[] procNames, XmlWriter lastProcessor)
        {
            XmlWriter currentProc = lastProcessor;
            if (procNames != null)
            {
                for (int i = procNames.Length - 1; i >= 0; --i)
                {
                    if (!Contains(procNames[i], this.skipedPostProcessors))
                    {
                        Type type = Type.GetType(procNames[i]);
                        object[] parameters = { currentProc };
                        XmlWriter newProc = (XmlWriter)Activator.CreateInstance(type, parameters);
                        currentProc = newProc;
                    }
                }
            }
            return currentProc;
        }

        private bool Contains(string processorFullName, ArrayList names)
        {
            foreach (string name in names)
            {
                if (processorFullName.Contains(name))
                {
                    return true;
                }
            }
            return false;
        }

        #region XSLTs events handling and redispatching
        /// <summary>
        /// Progress and Feedback listener functions
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e">contains the feedback message or null if the message is a progress event</param>
        public delegate void XSLTMessagesListener(object sender, EventArgs e);

        /// <summary>
        /// Progress messages redispatcher
        /// </summary>
        private event XSLTMessagesListener progressMessageIntercepted;
        /// <summary>
        /// Progress messages redispatcher
        /// </summary>
        private event XSLTMessagesListener progressMessageInterceptedMaster;
        /// <summary>
        /// 
        /// </summary>
        private event XSLTMessagesListener feedbackMessageIntercepted;
        /// <summary>
        /// 
        /// </summary>
        private event XSLTMessagesListener feedbackValidationIntercepted;

        /// <summary>
        /// Events handler that receives XSLT messages and redispatche them using the progress and feedback intercepted event launcher.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void onXSLTMessageEvent(object sender, XsltMessageEncounteredEventArgs e) {
            if (e.Message.StartsWith("progress:")) {
                if (progressMessageIntercepted != null) {
                    progressMessageIntercepted(this, null);
                }
            } else if (e.Message.StartsWith("translation.oox2Daisy.")) {
                fidilityLoss.Add(e.Message);

                if (feedbackMessageIntercepted != null) {
                    feedbackMessageIntercepted(this, new DaisyEventArgs(e.Message));

                }
            }
        }

        /// <summary>
        /// Secondary handler that only redispatch "progress" messages 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void onXSLTProgressMessageEvent(object sender, XsltMessageEncounteredEventArgs e) {
            if (e.Message.StartsWith("progress:")) {
                if (progressMessageInterceptedMaster != null) {
                    progressMessageInterceptedMaster(this, null);
                }
            }
        }

        

        public void AddProgressMessageListener(XSLTMessagesListener listener)
        {
            progressMessageIntercepted += listener;
        }

        public void AddProgressMessageListenerMaster(XSLTMessagesListener listener)
        {
            progressMessageInterceptedMaster += listener;
        }

        public void AddFeedbackMessageListener(XSLTMessagesListener listener)
        {
            feedbackMessageIntercepted += listener;
        }

        public void AddFeedbackValidationListener(XSLTMessagesListener listener)
        {
            feedbackValidationIntercepted += listener;
        }

        public void RemoveMessageListeners()
        {
            progressMessageIntercepted = null;
            feedbackMessageIntercepted = null;
            feedbackValidationIntercepted = null;
            progressMessageInterceptedMaster = null;
        }

        #endregion

        #region Files copy to output folder
        /* Function to Copy DTD file to the Output folder*/
        public void CopyDTDToDestinationfolder(String outputFile)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            Stream stream = null;
            string fileName = Path.GetDirectoryName(outputFile) + "\\dtbook-2005-3.dtd";
            foreach (string name in asm.GetManifestResourceNames())
            {
                if (name.EndsWith("dtbook-2005-3.dtd"))
                {
                    stream = asm.GetManifestResourceStream(name);
                    break;
                }

            }

            StreamReader reader = new StreamReader(stream);
            string data = reader.ReadToEnd();
            reader.Close();

            StreamWriter writer = new StreamWriter(fileName);
            writer.Write(data);
            writer.Close();

        }

        public void CopyCSSToDestinationfolder(String outputFile)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            Stream stream = null;
            string fileName = Path.GetDirectoryName(outputFile) + "\\dtbookbasic.css";
            foreach (string name in asm.GetManifestResourceNames())
            {
                if (name.EndsWith("dtbookbasic.css"))
                {
                    stream = asm.GetManifestResourceStream(name);
                    break;
                }

            }
            StreamReader reader = new StreamReader(stream);
            StreamWriter writer = new StreamWriter(fileName);
            string data = reader.ReadToEnd();
            writer.Write(data);
            reader.Close();
            writer.Close();
        }

        public void CopyMATHToDestinationfolder(String outputFile)
        {
            string fileName = "";
            fileName = Path.GetDirectoryName(outputFile) + "\\mathml2-qname-1.mod";
            CopyingAssemblyFile(fileName, "mathml2-qname-1.mod");
            fileName = Path.GetDirectoryName(outputFile) + "\\mathml2.DTD";
            CopyingAssemblyFile(fileName, "mathml2.DTD");

            for (int i = 0; i < MathEntities8879.Count; i++)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(outputFile) + "\\iso8879");
                fileName = Path.GetDirectoryName(outputFile) + "\\iso8879\\" + MathEntities8879[i].ToString();
                CopyingAssemblyFile(fileName, MathEntities8879[i].ToString());

            }

            for (int i = 0; i < MathEntities9573.Count; i++)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(outputFile) + "\\iso9573-13");
                fileName = Path.GetDirectoryName(outputFile) + "\\iso9573-13\\" + MathEntities9573[i].ToString();
                CopyingAssemblyFile(fileName, MathEntities9573[i].ToString());
            }

            for (int i = 0; i < MathMLEntities.Count; i++)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(outputFile) + "\\mathml");
                fileName = Path.GetDirectoryName(outputFile) + "\\mathml\\" + MathMLEntities[i].ToString();
                CopyingAssemblyFile(fileName, MathMLEntities[i].ToString());

            }
        }

        public void CopyingAssemblyFile(String destinationFile, String resourceFileName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            Stream stream = null;

            foreach (string name in asm.GetManifestResourceNames())
            {
                if (name.EndsWith(resourceFileName))
                {
                    stream = asm.GetManifestResourceStream(name);
                    break;
                }

            }

            StreamReader reader = new StreamReader(stream);
            StreamWriter writer = new StreamWriter(destinationFile);
            string data = reader.ReadToEnd();
            writer.Write(data);
            reader.Close();
            writer.Close();
        }
        #endregion

        #region XML validation
        /*Function to do validation of Output XML file with DTD*/
        public void XmlValidation(String outFile)
        {
            isValid = true;
            error = "";
            XmlTextReader xml = new XmlTextReader(outFile);
           
            XmlReaderSettings settings = new XmlReaderSettings();

            settings.ValidationType = ValidationType.DTD;
            settings.DtdProcessing = DtdProcessing.Parse;
            settings.ValidationEventHandler += new ValidationEventHandler(onValidationEvent);
            XmlReader xsd = XmlReader.Create(xml,settings);
            
            try
            {

                ArrayList errTxt = new ArrayList();
                for (int i = 0; i <= 4; i++)
                    errTxt.Add("");
                while (xsd.Read())
                {
                    errTxt[4] = errTxt[3];
                    errTxt[3] = errTxt[2];
                    errTxt[2] = errTxt[1];
                    errTxt[1] = errTxt[0];
                    errTxt[0] = xsd.ReadString();
                    errorText = "";
                    for (int i = 4; i >= 0; i--)
                        errorText = errorText + errTxt[i].ToString() + " ";
                    if (errorText.Contains("\n"))
                        errorText = errorText.Replace("\n", "");
                    if (errorText.Contains("\r"))
                        errorText = errorText.Replace("\r", "");
                    if (errorText.Length > 100)
                        errorText = errorText.Substring(0, 100);
                }
                xsd.Close();

                Stream stream = null;
                Assembly asm = Assembly.GetExecutingAssembly();
                foreach (string name in asm.GetManifestResourceNames())
                {
                    if (name.EndsWith("Schematron.xsl"))
                    {
                        stream = asm.GetManifestResourceStream(name);
                        break;
                    }
                }

                XmlReader rdr = XmlReader.Create(stream);
                XPathDocument doc = new XPathDocument(outFile);

                XslCompiledTransform trans = new XslCompiledTransform(true);
                trans.Load(rdr);

                XmlTextWriter myWriter = new XmlTextWriter(Path.GetDirectoryName(outFile) + "\\report.txt", null);
                trans.Transform(doc, null, myWriter);

                myWriter.Close();
                rdr.Close();

                StreamReader reader = new StreamReader(Path.GetDirectoryName(outFile) + "\\report.txt");
                if (!reader.EndOfStream)
                {
                    error += reader.ReadToEnd();

                    if (feedbackValidationIntercepted != null)
                    {
                        feedbackValidationIntercepted(this, new DaisyEventArgs(error));
                    }
                }
                reader.Close();

                if (File.Exists(Path.GetDirectoryName(outFile) + "\\report.txt"))
                {
                    File.Delete(Path.GetDirectoryName(outFile) + "\\report.txt");
                }

                // Check whether the document is valid or invalid.
                if (isValid == false)
                {
                    if (feedbackValidationIntercepted != null)
                    {
                        feedbackValidationIntercepted(this, new DaisyEventArgs(error));
                    }
                }
            }
            catch (UnauthorizedAccessException a)
            {
                xsd.Close();
                //dont have access permission
                error = a.Message;

                if (feedbackValidationIntercepted != null)
                {
                    feedbackValidationIntercepted(this, new DaisyEventArgs(error));
                }
            }
            catch (Exception a)
            {
                xsd.Close();
                //and other things that could go wrong
                error = a.Message;

                if (feedbackValidationIntercepted != null)
                {
                    feedbackValidationIntercepted(this, new DaisyEventArgs(error));
                }
            }
        }


        /// <summary>
        /// XML Validation events callback
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public void onValidationEvent(object sender, ValidationEventArgs args)
        {
            isValid = false;
            error += " Line Number : " + args.Exception.LineNumber + " and " +
             " Line Position : " + args.Exception.LinePosition + Environment.NewLine +
             " Message : " + args.Message + Environment.NewLine + " Reference Text :  " + errorText + Environment.NewLine;
        }

        #endregion XML validation

        #region Master and Subdocuments
        public void convertDocuments(ArrayList subList, string outputfilepath, ConversionParameters conversion) {
            try {
                int subDocFootNum = 1;
                string errors = "";

                XmlDocument mergeXmlDoc = new XmlDocument();
                ArrayList mergeDocLanguage = new ArrayList();
                ArrayList lostElements = new ArrayList();

                /*this.computeSize = true;
                converter.RemoveMessageListeners();
                converter.AddProgressMessageListenerMaster(new WordToDTBookXMLConverter.XSLTMessagesListener(onProgressMessageReceived));
                converter.AddFeedbackMessageListener(new WordToDTBookXMLConverter.XSLTMessagesListener(onFeedbackMessageReceived));

                for (int i = 0; i < subList.Count; i++) {
                    string[] splt = subList[i].ToString().Split('|');
                    docName = splt[0];

                    converter.Transform(splt[0], null, table, null, true, "");
                }*/

                string cutData = "";

                for (int i = 0; i < subList.Count; i++) {
                    string[] splt = subList[i].ToString().Split('|');
                    String outputFile = outputfilepath + "\\" + Path.GetFileNameWithoutExtension(splt[0]) + ".xml";
                    String ridOutputFile = splt[1];
                    string docName = splt[0];
                    this.convert(splt[0], outputFile, conversion, true, "Doc" + i);
                    ReplaceData(outputFile, out cutData);
                    if (i == 0) {
                        mergeXmlDoc.Load(outputFile);
                    } else {
                        MergeXml(outputFile, mergeXmlDoc, ridOutputFile, splt[0], ref subDocFootNum, ref mergeDocLanguage, ref lostElements, ref errors);
                    }
                    if (File.Exists(outputFile)) {
                        File.Delete(outputFile);
                    }
                }
                SetPageNum(mergeXmlDoc);
                SetImage(mergeXmlDoc);
                SetLanguage(mergeXmlDoc, mergeDocLanguage);
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
            } catch (Exception e) {
                error_Exception = manager.GetString("TranslationFailed") + "\n" + manager.GetString("WellDaisyFormat") + "\n" + " \"" + Path.GetFileName(tempInputFile) + "\"\n" + error_MasterSub + "\n" + "Problem is:" + "\n" + e.Message + "\n";

            }
        }

        public XmlDocument SetFootnote(XmlDocument mergeXmlDoc, String SubDocFootnum) {
            int footnoteCount = 1, endnoteCount = 1;
            XmlNodeList noteList = mergeXmlDoc.SelectNodes("//note");
            if (noteList != null) {
                for (int i = 1; i <= noteList.Count; i++) {
                    if (mergeXmlDoc.SelectNodes("//note").Item(i - 1).Attributes[1].InnerText == "Footnote") {
                        mergeXmlDoc.SelectNodes("//note").Item(i - 1).Attributes[0].InnerText = SubDocFootnum + "footnote-" + footnoteCount.ToString();
                        footnoteCount++;
                    }
                    if (mergeXmlDoc.SelectNodes("//note").Item(i - 1).Attributes[1].InnerText == "Endnote") {
                        mergeXmlDoc.SelectNodes("//note").Item(i - 1).Attributes[0].InnerText = SubDocFootnum + "endnote-" + endnoteCount.ToString();
                        endnoteCount++;
                    }
                }
            }

            footnoteCount = 1;
            endnoteCount = 1;
            noteList = mergeXmlDoc.SelectNodes("//noteref");
            if (noteList != null) {

                for (int i = 1; i <= noteList.Count; i++) {
                    if (mergeXmlDoc.SelectNodes("//noteref").Item(i - 1).Attributes[1].InnerText == "Footnote") {
                        mergeXmlDoc.SelectNodes("//noteref").Item(i - 1).Attributes[0].InnerText = "#" + SubDocFootnum + "footnote-" + footnoteCount.ToString();
                        footnoteCount++;
                    }
                    if (mergeXmlDoc.SelectNodes("//noteref").Item(i - 1).Attributes[1].InnerText == "Endnote") {
                        mergeXmlDoc.SelectNodes("//noteref").Item(i - 1).Attributes[0].InnerText = "#" + SubDocFootnum + "endnote-" + endnoteCount.ToString();
                        endnoteCount++;
                    }
                }
            }

            return mergeXmlDoc;

        }

        /* Function which creates unique ID to page numbers*/
        public void SetPageNum(XmlDocument mergeXmlDoc) {
            XmlNodeList pageList = mergeXmlDoc.SelectNodes("//pagenum");
            for (int i = 1; i <= pageList.Count; i++) {
                mergeXmlDoc.SelectNodes("//pagenum").Item(i - 1).Attributes[1].InnerText = "page" + i.ToString();
            }
        }

        /* Function which creates unique ID to Images*/
        public void SetImage(XmlDocument mergeXmlDoc) {
            XmlNodeList imageList = mergeXmlDoc.SelectNodes("//img");
            int j = 0;
            for (int i = 1; i <= imageList.Count; i++) {
                if (mergeXmlDoc.SelectNodes("//img").Item(i - 1).Attributes[0].InnerText.StartsWith("rId")) {
                    mergeXmlDoc.SelectNodes("//img").Item(i - 1).Attributes[0].InnerText = "rId" + j.ToString();
                    j++;
                }
            }
            XmlNodeList captionList = mergeXmlDoc.SelectNodes("//caption");
            for (int i = 1; i <= captionList.Count; i++) {
                XmlNode prevNode = mergeXmlDoc.SelectNodes("//caption").Item(i - 1).PreviousSibling;
                if (prevNode != null) {
                    String rId = prevNode.Attributes[0].InnerText;
                    mergeXmlDoc.SelectNodes("//caption").Item(i - 1).Attributes[0].InnerText = rId;
                }
            }
        }

        /* Function which creates language info of all sub documents in master.xml*/
        public void SetLanguage(XmlDocument mergeXmlDoc, ArrayList mergingLanguagesList) {
            XmlNodeList languageList = mergeXmlDoc.SelectNodes("//meta[@name='dc:Language']");

            for (int i = 0; i < languageList.Count; i++) {
                if (mergingLanguagesList.Contains(languageList[i].Attributes[1].Value)) {
                    int indx = mergingLanguagesList.IndexOf(languageList[i].Attributes[1].Value);
                    mergingLanguagesList.RemoveAt(indx);
                }
            }

            for (int i = 0; i < mergingLanguagesList.Count; i++) {
                XmlElement tempLang = mergeXmlDoc.CreateElement("meta");
                tempLang.SetAttribute("name", "dc:Language");
                tempLang.SetAttribute("content", mergingLanguagesList[i].ToString());
                mergeXmlDoc.SelectNodes("//head").Item(0).AppendChild(tempLang);
            }
        }

        /* Function which removes subdoc elements from the master.xml*/
        public void RemoveSubDoc(XmlDocument mergeXmlDoc) {
            XmlNodeList subDocList = mergeXmlDoc.SelectNodes("//subdoc");
            if (subDocList != null) {
                for (int i = 0; i < subDocList.Count; i++) {
                    subDocList.Item(i).ParentNode.RemoveChild(subDocList.Item(i));
                }
            }
        }



        public void MergeXml(
            string outputFile,
            XmlDocument mergeDoc,
            String rId,
            String inputFile,
            ref int subDocFootNum,
            ref ArrayList mergeDocLanguage,
            ref ArrayList lostElements,
            ref string errorsMessage
        ) {
            try {
                XmlNode tempNode = null;
                XmlDocument tempDoc = new XmlDocument();
                tempDoc.Load(outputFile);

                tempDoc = SetFootnote(tempDoc, "subDoc" + subDocFootNum);

                for (int i = 0; i < tempDoc.SelectSingleNode("//head").ChildNodes.Count; i++) {
                    tempNode = tempDoc.SelectSingleNode("//head").ChildNodes[i];

                    if (tempNode.Attributes[0].Value == "dc:Language") {
                        if (!mergeDocLanguage.Contains(tempNode.Attributes[1].Value)) {
                            mergeDocLanguage.Add(tempNode.Attributes[1].Value);
                        }
                    }
                }

                for (int i = 0; i < tempDoc.SelectSingleNode("//bodymatter").ChildNodes.Count; i++) {
                    tempNode = tempDoc.SelectSingleNode("//bodymatter").ChildNodes[i];

                    if (tempNode != null) {
                        XmlNode addBodyNode = mergeDoc.ImportNode(tempNode, true);
                        if (addBodyNode != null)
                            mergeDoc.SelectSingleNode("//subdoc[@rId='" + rId + "']").ParentNode.InsertBefore(addBodyNode, mergeDoc.SelectSingleNode("//subdoc[@rId='" + rId + "']"));
                    }
                }

                tempNode = tempDoc.SelectSingleNode("//frontmatter/level1[@class='print_toc']");
                if (tempNode != null) {
                    if (!lostElements.Contains("TOC is not translated" + " for " + Path.GetFileName(inputFile))) {
                        lostElements.Add("TOC is not translated" + " for " + Path.GetFileName(inputFile));
                    }

                }

                mergeDoc.SelectSingleNode("//subdoc[@rId='" + rId + "']").ParentNode.RemoveChild(mergeDoc.SelectSingleNode("//subdoc[@rId='" + rId + "']"));

                XmlNode node = tempDoc.SelectSingleNode("//rearmatter");

                if (node != null) {
                    for (int i = 0; i < tempDoc.SelectSingleNode("//rearmatter").ChildNodes.Count; i++) {
                        tempNode = tempDoc.SelectSingleNode("//rearmatter").ChildNodes[i];

                        if (tempNode != null) {
                            XmlNode addRearNode = mergeDoc.ImportNode(tempNode, true);
                            if (addRearNode != null)
                                mergeDoc.LastChild.LastChild.LastChild.AppendChild(addRearNode);
                        }
                    }
                }
                subDocFootNum++;
            } catch (Exception e) {
                errorsMessage = errorsMessage + "\n" + " \"" + inputFile + "\"";
                errorsMessage = errorsMessage + "\n" + "Validation error:" + "\n" + e.Message + "\n";
            }
        }

        /// <summary>
        /// See Progress ReplaceData method
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="cutData"></param>
        public void ReplaceData(String fileName, out string cutData) {
            StreamReader reader = new StreamReader(fileName);
            string data = reader.ReadToEnd();
            reader.Close();
            
            cutData = "";
            StreamWriter writer = new StreamWriter(fileName);
            bool hasMathMl = data.Contains("</mml:math>");
            string doctypeToRemove = !hasMathMl ?
                "<?xml-stylesheet href=\"dtbookbasic.css\" type=\"text/css\"?><!DOCTYPE dtbook PUBLIC '-//NISO//DTD dtbook 2005-3//EN' 'http://www.daisy.org/z3986/2005/dtbook-2005-3.dtd' >" :
                "<?xml-stylesheet href=\"dtbookbasic.css\" type=\"text/css\"?><!DOCTYPE dtbook PUBLIC '-//NISO//DTD dtbook 2005-3//EN' 'http://www.daisy.org/z3986/2005/dtbook-2005-3.dtd'[<!ENTITY % MATHML.prefixed \"INCLUDE\" ><!ENTITY % MATHML.prefix \"mml\"><!ENTITY % Schema.prefix \"sch\"><!ENTITY % XLINK.prefix \"xlp\"><!ENTITY % MATHML.Common.attrib \"xlink:href    CDATA       #IMPLIED xlink:type     CDATA       #IMPLIED   class          CDATA       #IMPLIED  style          CDATA       #IMPLIED  id             ID          #IMPLIED  xref           IDREF       #IMPLIED  other          CDATA       #IMPLIED   xmlns:dtbook   CDATA       #FIXED 'http://www.daisy.org/z3986/2005/dtbook/' dtbook:smilref CDATA       #IMPLIED\"><!ENTITY % mathML2 SYSTEM 'mathml2.dtd'>%mathML2;<!ENTITY % externalFlow \"| mml:math\"><!ENTITY % externalNamespaces \"xmlns:mml CDATA #FIXED 'http://www.w3.org/1998/Math/MathML'\">]>";
            // Remove doctype
            data = data.Replace(
                doctypeToRemove,
                "<?xml-stylesheet href=\"dtbookbasic.css\" type=\"text/css\"?>"
            );
            if (hasMathMl) { // I don't know what this part extracts exactly
                cutData = data.Substring(95, 1091);
                data = data.Remove(95, 1091);
            }

            // Remove namespace and add mathml if needed
            // (i assumed, but the previous codebase add the namespace
            // if no mml:math closing tag is found
            data = data.Replace(
                    "<dtbook xmlns=\"http://www.daisy.org/z3986/2005/dtbook/\" version=\"2005-3\"",
                    "<dtbook version=\"" + "2005-3\"" + (!hasMathMl ? " xmlns:mml=\"http://www.w3.org/1998/Math/MathML\"" : ""));

            writer.Write(data);
            writer.Close();
        }

        /* Function which merges subdocument.xml and master.xml*/
        public void ReplaceData(String fileName, bool value, in string cutData = "") {
            StreamReader reader = new StreamReader(fileName);
            string data = reader.ReadToEnd();
            reader.Close();
            string tempData = "";
            StreamWriter writer = new StreamWriter(fileName);
            if (value) {
                if (!data.Contains("</mml:math>")) {
                    data = data.Replace("<?xml-stylesheet href=\"dtbookbasic.css\" type=\"text/css\"?>", "<?xml-stylesheet href=\"dtbookbasic.css\" type=\"text/css\"?><!DOCTYPE dtbook SYSTEM 'dtbook-2005-3.dtd'>");
                    data = data.Replace("<dtbook version=\"" + "2005-3\" xmlns:mml=\"http://www.w3.org/1998/Math/MathML\" xml:lang=", "<dtbook version=\"" + "2005-3\" xml:lang=");
                } else {
                    tempData = cutData.Replace("<!DOCTYPE dtbook PUBLIC '-//NISO//DTD dtbook 2005-3//EN' 'http://www.daisy.org/z3986/2005/dtbook-2005-3.dtd'", "<!DOCTYPE dtbook SYSTEM 'dtbook-2005-3.dtd'");
                    tempData = tempData.Replace("<!ENTITY % mathML2 PUBLIC \"-//W3C//DTD MathML 2.0//EN\" \"http://www.w3.org/Math/DTD/mathml2/mathml2.dtd\">", "<!ENTITY % mathML2 SYSTEM 'mathml2.dtd'>");
                    data = data.Replace("<?xml-stylesheet href=\"dtbookbasic.css\" type=\"text/css\"?>", "<?xml-stylesheet href=\"dtbookbasic.css\" type=\"text/css\"?>" + tempData);
                }
            } else {
                if (!data.Contains("</mml:math>")) {
                    data = data.Replace("<!DOCTYPE dtbook SYSTEM 'dtbook-2005-3.dtd'>", "<!DOCTYPE dtbook PUBLIC '-//NISO//DTD dtbook 2005-3//EN' 'http://www.daisy.org/z3986/2005/dtbook-2005-3.dtd'>");
                    data = data.Replace("<dtbook version=\"" + "2005-3\"", "<dtbook xmlns=\"http://www.daisy.org/z3986/2005/dtbook/\" version=\"2005-3\"");
                } else {
                    data = data.Replace(tempData, cutData);
                    data = data.Replace("<dtbook version=\"" + "2005-3\"", "<dtbook xmlns=\"http://www.daisy.org/z3986/2005/dtbook/\" version=\"2005-3\"");
                }
            }
            writer.Write(data);
            writer.Close();
        }
        #endregion

    }


}
