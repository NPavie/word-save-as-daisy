using Daisy.SaveAsDAISY.DaisyConverterLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Daisy.SaveAsDAISY.CommandLineTool {
    class ConverterFactory {
        private static WordToDTBookXMLConverter wordInstance;

        protected ConverterFactory() {
        }

        public static WordToDTBookXMLConverter Instance(Direction transformDirection) {
            switch (transformDirection) {
                case Direction.DocxToXml:
                    if (wordInstance == null) {
                        wordInstance = new Daisy.SaveAsDAISY.Word.Converter();
                    }
                    return wordInstance;
                default:
                    throw new ArgumentException("invalid transform direction type");
            }
        }
    }
}
