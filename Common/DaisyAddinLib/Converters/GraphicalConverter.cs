using System;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;

using Daisy.SaveAsDAISY.Conversion;

namespace Daisy.SaveAsDAISY
{
	/// <summary>
	/// Converter extension using progress dialog and 
    /// - Request user confirmation p
    /// - Progression dialog : display progression message in a textarea instead of progress bar
    /// 
	/// </summary>
    public class GraphicalConverter : Converter
    {
        ConversionProgress progressDialog;

        
        public GraphicalConverter(GraphicalEventsHandler eventsHandler, WordToDTBookXMLTransform documentXMLConverter, ConversionParameters conversion) 
			: base(eventsHandler,documentXMLConverter, conversion)
        {
            progressDialog = new ConversionProgress();
            progressDialog.setCancelClickListener(this.requestConversionCancel);
            ((GraphicalEventsHandler) this.eventsHandler).LinkToProgressDialog(ref progressDialog);
        }



    }
}