﻿using Daisy.SaveAsDAISY.Conversion.Events;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2.Scripts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline.ChainedScripts {
    public class CleanedDtbookToEpub3 : Script {

        Script dtbookCleaner;
        Script dtbookToEpub3;

        public CleanedDtbookToEpub3(IConversionEventsHandler e): base(e) {
            this.niceName = "Export to EPUB3";
            dtbookCleaner = new DtbookCleaner(e);
            dtbookToEpub3 = new DtbookToEpub3(e);
            // TODO : for now we consider the 3 global steps of the progression but some granularity within
            // scripts could be taked in account
            StepsCount = 4;
            // set dtbook cleaner to apply default cleanups
            dtbookCleaner.Parameters["tidy"].ParameterValue = true;
            dtbookCleaner.Parameters["repair"].ParameterValue = true;
            dtbookCleaner.Parameters["narrator"].ParameterValue = true;
            // use dtbook to epub3 parameters
            _parameters = new Dictionary<string, ScriptParameter>
            {
                { "input", new ScriptParameter(
                        "source",
                        "OPF",
                        new PathDataType(
                            PathDataType.InputOrOutput.input,
                            PathDataType.FileOrDirectory.File,
                            "application/oebps-package+xml"
                        ),
                        "",
                        true,
                        "The package file of the input DTB.",
                        true,
                        ScriptParameter.ParameterDirection.Input
                    )
                },
                //{"tts-log", new ScriptParameter(
                //        "tts-log",
                //        "TTS log output directory",
                //        new PathDataType(
                //            PathDataType.InputOrOutput.output,
                //            PathDataType.FileOrDirectory.Directory
                //        ),
                //        "",
                //        false,
                //        "TTS log output directory",
                //         true,
                //        ScriptParameter.ParameterDirection.Output
                //    )
                //},
                {"language", new ScriptParameter(
                        "language",
                        "Language code",
                        new StringDataType(),
                        "",
                        false,
                        "Language code of the input document."
                    )
                },
                {"output", new ScriptParameter(
                        "output-dir",
                        "EPUB output directory",
                        new PathDataType(
                            PathDataType.InputOrOutput.output,
                            PathDataType.FileOrDirectory.Directory
                        ),
                        "",
                        true,
                        "The produced EPUB."
                    )
                },
                { "validation", new ScriptParameter(
                        "validation",
                        "Validation",
                        new EnumDataType(
                            new Dictionary<string, object> {
                                { "No validation", "off" },
                                { "Report validation issues", "report" },
                                { "Abort on validation issues", "abort" },
                            }, "Abort on validation issues"),
                        "abort",
                        false,
                        "Whether to abort on validation issues."
                    )
                },
                { "validation-report", new ScriptParameter(
                        "validation-report",
                        "Validation reports",
                        new PathDataType(
                            PathDataType.InputOrOutput.output,
                            PathDataType.FileOrDirectory.Directory
                        ),
                        "",
                        false,
                        "Output path of the validation reports",
                        false,
                        ScriptParameter.ParameterDirection.Output
                    )
                },
                //{ "validation-status", new ScriptParameter(
                //        "validation-report",
                //        "Validation reports",
                //        new PathDataType(
                //            PathDataType.InputOrOutput.output,
                //            PathDataType.FileOrDirectory.Directory
                //        ),
                //        "",
                //        false,
                //        ""
                //    )
                //},
                {"nimas", new ScriptParameter(
                        "nimas",
                        "NIMAS input",
                        new BoolDataType(),
                        false,
                        false,
                        "Whether the input DTBook is a NIMAS 1.1-conformant XML content file.",
                        false // not sure this option should available in save as daisy
                    )
                },
                {"tts-config", new ScriptParameter(
                        "tts-config",
                        "Text-to-speech configuration file",
                        new PathDataType(PathDataType.InputOrOutput.input,PathDataType.FileOrDirectory.File),
                        "",
                        false,
                        "Configuration file for the text-to-speech.\r\n\r\n[More details on the configuration file format](http://daisy.github.io/pipeline/Get-Help/User-Guide/Text-To-Speech/).",
                        true,
                        ScriptParameter.ParameterDirection.Input
                    )
                },
                {"audio", new ScriptParameter(
                        "audio",
                        "Enable text-to-speech",
                        new BoolDataType(),
                        false,
                        false,
                        "Whether to use a speech synthesizer to produce audio files."
                    )
                },
            };
        }

        public override void ExecuteScript(string inputPath, bool isQuite) {

            // Create a directory using the document name
            string finalOutput = Path.Combine(
                Parameters["output"].ParameterValue.ToString(),
                Path.GetFileNameWithoutExtension(inputPath)
                ) ;
            // Remove and recreate result folder
            // Since the DaisyToEpub3 requires output folder to be empty
            if (Directory.Exists(finalOutput)) {
                Directory.Delete(finalOutput, true);
            }
            Directory.CreateDirectory(finalOutput);


            DirectoryInfo tempDir = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), Path.GetRandomFileName()));
#if DEBUG
            this.EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Cleaning " + this._parameters["input"].ParameterValue + " into " + tempDir.FullName));
#else
            this.EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Cleaning the DTBook XML... "));
#endif
            dtbookCleaner.Parameters["input"].ParameterValue = this._parameters["input"].ParameterValue;
            dtbookCleaner.Parameters["output"].ParameterValue = tempDir.FullName;
            dtbookCleaner.ExecuteScript(inputPath,true);

            // transfer parameters value
            foreach(var kv in this._parameters)
            {
                if (dtbookToEpub3.Parameters.ContainsKey(kv.Key))
                {
                    dtbookToEpub3.Parameters[kv.Key] = kv.Value;
                }
            }
            // rebind input and output
            dtbookToEpub3.Parameters["input"].ParameterValue = Directory.GetFiles(tempDir.FullName, "*.xml")[0];
            dtbookToEpub3.Parameters["output"].ParameterValue = Directory.CreateDirectory(Path.Combine(finalOutput, "EPUB3")).FullName;
            dtbookToEpub3.Parameters["validation-report"].ParameterValue = Directory.CreateDirectory(Path.Combine(finalOutput, "report")).FullName;
            dtbookToEpub3.Parameters["tts-log"].ParameterValue = Directory.CreateDirectory(Path.Combine(finalOutput, "tts-log")).FullName;
#if DEBUG
            this.EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Converting " + dtbookToEpub3.Parameters["input"].ParameterValue + " dtbook XML to EPUB3 in " + dtbookToEpub3.Parameters["output"].ParameterValue));
#else
            this.EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Converting DTBook XML to EPUB3..."));
#endif

            dtbookToEpub3.ExecuteScript(dtbookToEpub3.Parameters["input"].ParameterValue.ToString());

            //if (File.Exists(Path.Combine(finalOutput, "result.epub"))) {
            //    this.EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Moving result.epub to " + Path.Combine(finalOutput, Path.GetFileNameWithoutExtension(inputPath) + ".epub")));
                
            //    File.Move(
            //        Path.Combine(finalOutput, "result.epub"),
            //        Path.Combine(finalOutput, Path.GetFileNameWithoutExtension(inputPath) + ".epub")
            //    );
            //    this.EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Successfully converted the document to " + Path.Combine(finalOutput, Path.GetFileNameWithoutExtension(inputPath) + ".epub")));

            //}
        }
    }
}
