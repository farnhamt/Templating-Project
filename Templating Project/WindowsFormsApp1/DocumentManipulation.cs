using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace TemplatingProject {
	public class DocumentManipulation {
		/// <summary>Color pallette that is used on the graphs. Set by default, by can be changed by a user document command
		public System.Drawing.Color[] colorPalette = {System.Drawing.Color.FromArgb(215, 63, 9), System.Drawing.Color.FromArgb(170, 157, 46), System.Drawing.Color.FromArgb(183, 169, 154)};
		/// <summary>
		/// Creates a new Word.Application from the document that was selected by the user
		/// </summary>
		public Word.Application OpenDocument(string filepath) {
			Word.Application wordApp = new Word.Application();
			try {
				wordApp.Application.Documents.Open(filepath, false, false);
			}
			catch (Exception error) {
				MessageBox.Show("Failed to open document!\nError: " + error.Message + "\nTry closing all Microsoft Word processes from Task Manager.");
				Environment.Exit(1);
				return null;
			}
			return wordApp;
		}
		/// <summary>
		/// Finds and replaces a string of text with another string of text in the given word document
		/// </summary>
		private void ReplaceTextWithText(string textToReplace, string replacementText, Word.Application wordApp) {
			object replaceAll = Word.WdReplace.wdReplaceAll;
			Word.Find findObject = wordApp.Selection.Find;
			findObject.Execute(textToReplace, true, true, false, false, false, false, Word.WdFindWrap.wdFindAsk, false, replacementText, replaceAll, false, false, false, false);
		}
		/// <summary>
		/// Finds and replaces a string of text with an image specified at the given image file path
		/// </summary>
		private void ReplaceTextWithImage(string rawCommand, string imagePath, Word.Application wordApp) {
			Word.Selection sel = wordApp.Selection;
			string keyword = rawCommand;
			try {
				sel.Find.Execute(keyword, Replace: Word.WdReplace.wdReplaceOne);
			}
			catch(Exception e) {
				string errorMessage = "Error parsing document command at command:\n" + keyword;
				if (keyword.Length > 250) {
					errorMessage += "\n\nCommand exceeds maximum length";
				}
				MessageBox.Show(errorMessage);
				wordApp.Quit();
				Environment.Exit(1);
			}
			sel.Range.Select();
			var imagePath1 = Path.GetFullPath(string.Format(imagePath, keyword));
			sel.InlineShapes.AddPicture(imagePath, false, true);
		}
		/// <summary>
		/// Prompts the user to save the document as 
		/// </summary>
		private void SaveDocumentAs(string filename, Word.Application wordApp) {
			try {
				SaveFileDialog saveFileDialog = new SaveFileDialog();
				saveFileDialog.AutoUpgradeEnabled = false;
				if (saveFileDialog.ShowDialog() == DialogResult.OK) {
					wordApp.Application.ActiveDocument.SaveAs2(saveFileDialog.FileName);

					//wordApp.Application.ActiveDocument.SaveAs2(@"C:\VSTesting\" + filename + ".docx");
					//wordApp.Application.ActiveDocument.SaveAs(@"C:\VSTesting\" + filename + ".pdf", Word.WdSaveFormat.wdFormatPDF);
				}
			}
			catch (Exception e) {
				DialogResult result = MessageBox.Show("Failed to open save file dialog.\nTry again?", "Error:", MessageBoxButtons.YesNo);
				if (result == DialogResult.Yes) {
					SaveDocumentAs(filename, wordApp);
					return;
				}
				else {
					wordApp.Quit();
					return;
				}
			}
			wordApp.Quit();
		}

		public TextReplacementOptions GetTextReplacementOptions(string rawText) {
			TextReplacementOptions textReplacementOptions = new TextReplacementOptions();
			textReplacementOptions.rawInput = rawText;
			string[] options;
			string outputType;
			string outputOption1;
			List<string> columnNames = new List<string>();
			try {
				options = rawText.Split(';');
				outputType = options[0].ToLower();
				outputOption1 = options[1].ToLower();
			}
			catch (Exception e) {
				MessageBox.Show("Error in word document input at:\n" + rawText);
				textReplacementOptions.hasFailed = true;
				return textReplacementOptions;
			}
			if (outputType.Contains("colors") || outputType.Contains("colorpalette")) {
				System.Drawing.Color[] colorPallete = new System.Drawing.Color[options.Length - 1];
				for (int i = 0; i < options.Length - 1; i++) {
					string[] rgb = options[i + 1].Split(',');
					int[] rgbInts = new int[3];
					for (int j = 0; j < rgb.Length; j++) {
						rgbInts[j] = Convert.ToInt32(rgb[j].Trim(' ', '}'));
					}
					colorPallete[i] = System.Drawing.Color.FromArgb(rgbInts[0], rgbInts[1], rgbInts[2]);
				}
				this.colorPalette = colorPallete;
				textReplacementOptions.isColors = true;
				return textReplacementOptions;
			}
			if (outputType.Contains("bar")) {
				textReplacementOptions.isGraph = true;
				textReplacementOptions.graphType = "bar";
			}
			else if (outputType.Contains("pie")) {
				textReplacementOptions.isGraph = true;
				textReplacementOptions.graphType = "pie";
			}
			else {
				textReplacementOptions.isText = true;
			}
			if (outputOption1.Contains("range")) {
				textReplacementOptions.isRange = true;
			}
			else if (outputOption1.Contains("mean")) {
				textReplacementOptions.isMean = true;
			}
			else if (outputOption1.Contains("percentage") || outputOption1.Contains("%")) {
				textReplacementOptions.isPercentage = true;
			}
			else if (outputOption1.Contains("count")) {
				textReplacementOptions.isCount = true;
			}
			else {
				textReplacementOptions.isColumnValue = true;
			}
			if (textReplacementOptions.isGraph) {
				int fontSize = 0;
				int optionsIndex = 2;
				if (int.TryParse(options[2].Trim(' ', '}'), out fontSize)) {
					optionsIndex++;
				}
				textReplacementOptions.fontSize = fontSize;
				for (int i = optionsIndex; i < options.Length - 1; i++) {
					columnNames.Add(options[i].Trim(' ', '}'));
				}
				textReplacementOptions.graphTitle = options[options.Length - 1].Trim(' ', '}');
			}
			else if(textReplacementOptions.isColumnValue){
				columnNames.Add(outputOption1.Trim(' ', '}'));
			}
			else {
				for (int i = 2; i < options.Length; i++) {
					columnNames.Add(options[i].Trim(' ', '}'));
				}
			}
			textReplacementOptions.columnNames = columnNames;
			return textReplacementOptions;
		}


		private string GetCommand(Word.Application wordApp) {
			Word.Find findObject = wordApp.Selection.Find;
			findObject.ClearFormatting();
			findObject.Replacement.ClearFormatting();
			findObject.Forward = true;
			findObject.Wrap = Word.WdFindWrap.wdFindContinue;

			//Decalare a selection of our word document
			Word.Selection sel;
			try {
				sel = wordApp.Selection;
				int i = 0;
				//set the keyword as the text to find in the document
				sel.Find.Text = "{{{";
				if (!sel.Find.Execute()) {
					return "EOF";
				}
				Word.Range range = sel.Range;

				//Selects all of the text to replace including the curly braces.
				while (!sel.Text.Contains("}}}")) {

					sel.SetRange(range.Start, range.End + i);
					i++;
				}
				return sel.Text;
			}
			catch (Exception e) {
				MessageBox.Show("Error: Failed to find text to replace in document");
				return "error";
			}
		}
		/// <summary>
		/// Takes a command provided by the user in the word document, parses it into a TextReplacementOptions object, and executes that command based on the user-provided parameters.
		/// </summary>
		/// <param name="rawCommand">The exact text provided by the user in the document command</param>
		/// <param name="columnValueCounters">List of objects that contain information about each column</param>
		private void ProcessDocumentCommand(string rawCommand, Word.Application wordApp, List<ColumnValueCounter> columnValueCounters) {
			//Remove any extraneous braces
			string command = rawCommand.Trim('{', '}');
			//Parse the raw command to get a textReplacementOptions object that stores more usable command options.
			TextReplacementOptions processedCommand = GetTextReplacementOptions(command);
			//A list of the columns that we actually use for this command.
			List<ColumnValueCounter> usedColumns = new List<ColumnValueCounter>();
			DocumentCommandExecuter commandExecuter = new DocumentCommandExecuter(colorPalette);
			//if the command was a declaration of the color pallette used in the document, just remove the text and do not execute anything.
			if (processedCommand.isColors) {
				ReplaceTextWithText(rawCommand, "", wordApp);
				return;
			}
			//Populate a list of columnValueCounters that pertain to the columns that we are actually using for this command.
			for (int i = 0; i < columnValueCounters.Count; i++) {
				if (processedCommand.columnNames.Contains(columnValueCounters[i].columnName.ToLower()) 
					|| processedCommand.columnNames.Contains(columnValueCounters[i].columnName)
					|| processedCommand.columnNames.Contains(columnValueCounters[i].abbreviatedRepresentation)) {
					usedColumns.Add(columnValueCounters[i]);
				}
			}
			//Ensure that there are the same number of unique row values in each column value counter.
			int uniqueRowValueCount = 0;
			ColumnValueCounter maxColumn = new ColumnValueCounter();
			foreach (ColumnValueCounter column in usedColumns) {
				if (column.uniqueRowValues.Count > uniqueRowValueCount) {
					uniqueRowValueCount = column.uniqueRowValues.Count;
					maxColumn = column;
				}
			}
			if (usedColumns.Count > 1) {
				//For each column value counter
				for (int i = 0; i < usedColumns.Count; i++) {
					ColumnValueCounter currentColumn = usedColumns[i];
					//If the number of unique row values of the current column is less than any of the others
					if (currentColumn.uniqueRowValues.Count < uniqueRowValueCount) {
						//for each intended unique row value
						for (int j = 0; j < maxColumn.uniqueRowValues.Count; j++) {
							//for each unique row value in the column with missing unique row values
							for (int k = 0; k < maxColumn.uniqueRowValues.Count; k++) {
								//check to see if the unique row value that we are currently looking at from the list 
								//of inteded unique row values is also in the column with the list of unique row values that is missing some;
								if (currentColumn.uniqueRowValues[k].name == maxColumn.uniqueRowValues[j].name) {
									break;
								}
								//If we did not find the unique row value in this column then add it to the column
								if (k == (currentColumn.uniqueRowValues.Count - 1)) {
									List<UniqueRowValue> tempUniqueRowValues = new List<UniqueRowValue>();
									for (int l = 0; l < j; l++) {
										tempUniqueRowValues.Add(currentColumn.uniqueRowValues[l]);
									}
									tempUniqueRowValues.Add(new UniqueRowValue(maxColumn.uniqueRowValues[j].name, 0));
									for (int l = j; l < currentColumn.uniqueRowValues.Count; l++) {
										tempUniqueRowValues.Add(currentColumn.uniqueRowValues[l]);
									}
									currentColumn.uniqueRowValues = tempUniqueRowValues;
									break;
								}
							}
						}
					}
				}
			}
			if (processedCommand.hasFailed) {
				MessageBox.Show("Command Failed: Check template command syntax");
				wordApp.Quit();
			}
			//Process the command as either a graph or text replacement.
			if (processedCommand.isGraph) {
				string replaceWith = commandExecuter.GenerateGraph(usedColumns, Application.StartupPath + "\\tempGraph.PNG", processedCommand, wordApp);
				ReplaceTextWithImage(rawCommand, replaceWith, wordApp);
			}
			else if (processedCommand.isText) {
				string replaceWith = commandExecuter.GenerateText(usedColumns, rawCommand, processedCommand, wordApp);
				ReplaceTextWithText(rawCommand, replaceWith, wordApp);
			}
			else {
				return;
			}
		}

		public void ProcessDocument(Word.Application wordApp, List<ColumnValueCounter> columnValueCounters){
			string documentCommand = "";
			//Keep getting commands and processing them until there are no more commands to find.
			while (true) { 
				documentCommand = GetCommand(wordApp);
				if (documentCommand == "EOF") {
					break;
				}
				else if (documentCommand == "error") {
					wordApp.Quit();
					return;
				}
				else {
					ProcessDocumentCommand(documentCommand, wordApp, columnValueCounters);
				}

			}
			SaveDocumentAs("helloWorld", wordApp);
		}
		/// <summary>
		/// Struct that holds the color pallette, a list of column names, and 
		/// all of the other information that is relavent to deciding what to do with a user command.
		/// </summary>
		public struct TextReplacementOptions {
			/// <summary>Determines whether a command is declaring a color pallette for the document or not.
			public bool isColors;
			//set of descriptive booleans that determine how the data will be presented.
			public bool isGraph;
			public bool isText;
			public bool isCount;
			public bool isPercentage;
			public bool isMean;
			public bool isRange;
			public bool isColumnValue;
			/// <summary>Set when a command fails to be parsed into one of these objects
			public bool hasFailed;
			/// <summary>x-axis label font size
			public int fontSize;
			public string graphType;
			public string graphTitle;
			public string rawInput;
			public List<string> columnNames;
			public System.Drawing.Color[] colorPalette;
		}
	}

}
