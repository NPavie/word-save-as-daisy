using System.Xml;

using System.Windows.Forms;
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Text;
using System.Globalization;

using Microsoft.Office.Interop.Word;

using MSWordApp = Microsoft.Office.Interop.Word.Application;
using Daisy.SaveAsDAISY.Addins.Word2007;
using Daisy.SaveAsDAISY.Conversion;
using Daisy.SaveAsDAISY.Conversion.Events;
using Daisy.SaveAsDAISY.Forms;

namespace Daisy.SaveAsDAISY
{
	public partial class Launcher : Form
	{
		private string inputPath = "";

		public Launcher()
		{
			InitializeComponent();

			if (inputPath == "")
			{
				StartConversion.Enabled = false;
			}
		}

		private void CopyDirectory(string inputDirPath, string outputDirPath, bool overwrite = true)
        {
            if (!Directory.Exists(inputDirPath))
            {
				throw new ArgumentException("Directory does not exists", "inputDirPath");
            }
            if (!Directory.Exists(outputDirPath))
            {
				Directory.CreateDirectory(outputDirPath);
			}
			foreach(string dir in Directory.GetDirectories(inputDirPath,"*", SearchOption.AllDirectories))
            {
				Directory.CreateDirectory(dir.Replace(inputDirPath,outputDirPath));
			}
            foreach (string file in Directory.GetFiles(inputDirPath,"*",SearchOption.AllDirectories))
            {
				File.Copy(file, file.Replace(inputDirPath, outputDirPath), overwrite);
            }
        }

		/// <summary>
		/// Normalized string with removed diacritics, and spaces replaced by _
		/// </summary>
		/// <param name="stIn"></param>
		/// <returns></returns>
		static string Normalize(string stIn)
        {

			string stFormD = stIn.Normalize(NormalizationForm.FormD);
			StringBuilder sb = new StringBuilder();

			for (int ich = 0; ich < stFormD.Length; ich++)
			{
				UnicodeCategory uc = CharUnicodeInfo.GetUnicodeCategory(stFormD[ich]);
				if (uc != UnicodeCategory.NonSpacingMark)
				{
					sb.Append(stFormD[ich]);
				}
			}

			return (sb.ToString().Normalize(NormalizationForm.FormC));

		}


		private void Launcher_OnDragDrop(object sender , DragEventArgs e)
		{
			string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
			if (files != null && files.Length > 0)
			{
				 selectedFilePath.Text = inputPath = files[0];
			}
		}

		private void Launcher_OnDragEnter(object sender , DragEventArgs e)
		{
			if (e.Data.GetDataPresent(DataFormats.FileDrop))
			{
				e.Effect = DragDropEffects.Copy;
			}
			else e.Effect = DragDropEffects.None;
		}

		/// <summary>
		/// Bouton de lancement du traitement
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void launchScript_Click(object sender , EventArgs e)
		{
			if(inputPath == "")
			{
				MessageBox.Show("Please select a document to convert", "No document selected" , MessageBoxButtons.OK , MessageBoxIcon.Error);
			} else if (!File.Exists(inputPath))
			{
				MessageBox.Show(
					"Le dossier \r\n"
						+ inputPath + "\r\n" 
						+ "est introuvable.\r\n" 
						+ "Veuillez le déplacer sur votre poste ou sélectionnez un autre fichier"
					, "DOssier introuvable"
					, MessageBoxButtons.OK
					, MessageBoxIcon.Error
				);
			} else 
			{
				Dictionary<string, Match> workfolderFound = new Dictionary<string, Match>();
				Regex expectedStructure = new Regex(
						@"(\d{13})[\s]+(.*)"
						 , RegexOptions.Compiled | RegexOptions.IgnoreCase
					);
				if (inputPath.EndsWith("/") || inputPath.EndsWith("\\"))
				{
					inputPath = inputPath.Substring(0, inputPath.Length - 1);
				}
				

				try
				{
					MSWordApp applicationObject = new MSWordApp();
					//applicationObject.Visible = true;

					IDocumentPreprocessor preprocess = new DocumentPreprocessor(applicationObject);
					IConversionEventsHandler eventsHandler = null;

					WordToDTBookXMLTransform documentConverter = new WordToDTBookXMLTransform();
					Converter converter = null;
					eventsHandler = new GraphicalEventsHandler();
					ConversionParameters conversion = new ConversionParameters(applicationObject.Version);
					conversion.Visible = false;
					converter = new GraphicalConverter(preprocess, documentConverter, conversion, (GraphicalEventsHandler)eventsHandler);
					
					DocumentParameters currentDocument = converter.PreprocessDocument(inputPath);
                    if (((GraphicalConverter)converter).requestUserParameters(currentDocument) == ConversionStatus.ReadyForConversion)
                    {
                        ConversionResult result = converter.Convert(currentDocument);

                    }
                    else
                    {
                        eventsHandler.onConversionCanceled();
                    }
#if !DEBUG	
					MessageBox.Show("Conversion done");
#endif

                    applicationObject.Quit();

                } catch (Exception ex){
					string test = ex.Message;
					MessageBox.Show(
						"An error occured during conversion :\r\n"
							+ ex.Message
							+ "\r\n" 
							+ (ex.StackTrace != null 
								? ( "Stack trace:\r\n" 
									+ ex.StackTrace
								) : ""
							)
						, "An error occured"
						, MessageBoxButtons.OK
						, MessageBoxIcon.Error
					); 
				}

			}
		}

		private void browseFile_Click(object sender , EventArgs e)
		{
			using (OpenFileDialog openFileDialog = new OpenFileDialog())
			{
				openFileDialog.InitialDirectory = "c:\\";
				openFileDialog.Filter = "Word document (*.docx)|*.docx|Tous les fichiers (*.*)|*.*";
				if (openFileDialog.ShowDialog() == DialogResult.OK)
				{
					inputPath = openFileDialog.FileName;
					selectedFilePath.Text = inputPath;
					StartConversion.Enabled = true;
				}
			}
		}

		private void selectedFilePath_TextChanged(object sender , EventArgs e)
		{
			inputPath = selectedFilePath.Text;
			if(inputPath.StartsWith("\"") && inputPath.EndsWith("\""))
			{
				inputPath = inputPath.Substring(1 , inputPath.Length - 2);
			}
			StartConversion.Enabled = (inputPath != "");
		}
	}
}
