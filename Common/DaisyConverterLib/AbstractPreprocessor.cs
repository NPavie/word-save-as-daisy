using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Daisy.SaveAsDAISY.DaisyConverterLib {
    public interface IConversionPreprocessor {

        PreprocessingData preprocessDocument(
            IConversionEventsHandler eventsHandler,
            object documentToPreprocess,
            Pipeline postprocessingPipeline,
            string conversionMode
        );


    }
}
