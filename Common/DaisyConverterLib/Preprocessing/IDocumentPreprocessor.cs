using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Daisy.SaveAsDAISY.Conversion {

    public enum PreProcessingStatus {
        None, // No preprocessing done yet
        ValidatedName, // Name validation process
        CreatedWorkingCopy, // Making the working copy
        ProcessedShapes,
        ProcessedInlineShapes,
        ProcessedMathML,
        Canceled,
        Error,
        Success
    }

    /// <summary>
    /// Document preprocessing requires function
    /// </summary>
    public interface IDocumentPreprocessor {

        object startPreprocessing(DocumentParameters document, Events.IConversionEventsHandler eventsHandler = null);

        PreProcessingStatus ValidateName(ref object preprocessedObject, FilenameValidator authorizedNamePattern, Events.IConversionEventsHandler eventsHandler = null);

        PreProcessingStatus CreateWorkingCopy(ref object preprocessedObject, DocumentParameters document, Events.IConversionEventsHandler eventsHandler = null);

        PreProcessingStatus ProcessShapes(ref object preprocessedObject, DocumentParameters document, Events.IConversionEventsHandler eventsHandler = null);

        PreProcessingStatus ProcessInlineShapes(ref object preprocessedObject, DocumentParameters document, Events.IConversionEventsHandler eventsHandler = null);

        PreProcessingStatus ProcessEquations(ref object preprocessedObject, DocumentParameters document, Events.IConversionEventsHandler eventsHandler = null);

        PreProcessingStatus endPreprocessing(ref object preprocessedObject, Events.IConversionEventsHandler eventsHandler = null);

    }
}
