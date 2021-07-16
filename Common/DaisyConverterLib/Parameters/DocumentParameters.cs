using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Daisy.SaveAsDAISY.DaisyConverterLib {

    /// <summary>
    /// Document specific parameters, to be extracted by Preprocessors
    /// </summary>
    public class DocumentParameters {

        /// <summary>
        /// Word document type between :<br/>
        /// - Simple : the document is self contained<br/>
        /// - Master : the document refers to subdocuments<br/>
        /// - Sub : the document is refered by another document
        /// </summary>
        /*public enum DocType {
            Simple,
            Master,
            Sub
        }*/

        /// <summary>
        /// Original path/URL of the input
        /// </summary>
        public string InputPath { get; set; }

        /// <summary>
        ///
        /// </summary>
        public string OutputPath { get; set; }

        /// <summary>
        /// Temporary copy of the original input file
        /// </summary>
        public string TempInputFile { get; set; }

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
    }
}
