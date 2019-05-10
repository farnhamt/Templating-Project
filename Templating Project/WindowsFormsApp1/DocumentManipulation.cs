using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace TemplatingProject {
	/// <summary>
	/// This class handles all interactions with word documents. It contains functions for finding and replacing text in a word document with either different text or an image,
	/// finding and parsing commands that are given by the user within the word document, and uses a DocumentCommandExecuter to execute the parsed commands.
	/// </summary>
	public class DocumentManipulation {
		#region Class Wide Variables
		/// <summary>Color pallette that is used on the graphs. Set by default, by can be changed by a user document command</summary>
		public System.Drawing.Color[] ColorPalette = {System.Drawing.Color.FromArgb(215, 63, 9), System.Drawing.Color.FromArgb(170, 157, 46), System.Drawing.Color.FromArgb(183, 169, 154)};
		/// <summary>
		/// List of strings that represent unique row value names. The order of the strings in this list is the order of their corresponding graph columns and legend entries.
		/// For example, if this list stores ['Yes', 'No'], each graph will place the 'Yes' row at the first index of the graph and the 'No' row at the second index of the graph if the graph is using these row values.
		/// </summary>
		private List<string> _graphColumnOrder;
		#endregion
		#region OpenDocument
		/// <summary>
		/// Creates a new Word.Application from the document that was selected by the user
		/// </summary>
		public Word.Application OpenDocument(string filepath) {
			Word.Application wordApp = new Word.Application();
			try {
				wordApp.Application.Documents.Open(filepath, false, false);
			}
			catch (Exception error) {
				MessageBox.Show(new Form { TopMost = true }, "Failed to open document!\nError: " + error.Message + "\nTry closing all Microsoft Word processes from Task Manager.");
				Environment.Exit(1);
				return null;
			}
			return wordApp;
		}
		#endregion
		#region Text Replacement Functions
		/// <summary>
		/// Finds and replaces a string of text with another string of text in the given word document
		/// </summary>
		private void ReplaceTextWithText(string textToReplace, string replacementText, Word.Application wordApp) {
			object replaceAll = Word.WdReplace.wdReplaceAll;
			Word.Find findObject = wordApp.Selection.Find;
			try {
				//Attempt to find and replace the given text to replace with the given replacement text in one command.
				findObject.Execute(textToReplace, true, true, false, false, false, false, Word.WdFindWrap.wdFindAsk, false, replacementText, replaceAll, false, false, false, false);
			}
			catch (Exception e) {
				MessageBox.Show(new Form { TopMost = true }, e.Message + "at command: " + textToReplace);
				wordApp.Quit();
				Environment.Exit(1);
			}
		}
		/// <summary>
		/// Finds and replaces a string of text with an image specified at the given image file path
		/// </summary>
		private void ReplaceTextWithImage(string rawCommand, string imagePath, Word.Application wordApp) {
			Word.Selection sel = wordApp.Selection;
			string keyword = rawCommand;
			try {
				//Attempt to find the text that will be replaced.
				sel.Find.Execute(keyword, Replace: Word.WdReplace.wdReplaceOne);
			}
			catch(Exception) {
				string errorMessage = "Error parsing document command at command:\n" + keyword;
				if (keyword.Length > 250) {
					errorMessage += "\n\nCommand exceeds maximum length";
				}
				MessageBox.Show(new Form { TopMost = true }, errorMessage);
				wordApp.Quit();
				Environment.Exit(1);
			}
			//Select the text that was found.
			sel.Range.Select();
			//Get the full file path of the given image.
			var imagePath1 = Path.GetFullPath(string.Format(imagePath, keyword));
			//Replace the text with the image at the given path.
			sel.InlineShapes.AddPicture(imagePath, false, true);
		}
		#endregion
		#region SaveDocumentAs
		/// <summary>
		/// Prompts the user to save the document as 
		/// </summary>
		private void SaveDocumentAs(string filename, Word.Application wordApp) {
			try {
				//Create a topmost form and show the file dialog on that form.
				Form topmostForm = new Form { TopMost = true };
				SaveFileDialog saveFileDialog = new SaveFileDialog { AutoUpgradeEnabled = false	};
				if (saveFileDialog.ShowDialog(topmostForm) == DialogResult.OK) {
					wordApp.Application.ActiveDocument.SaveAs2(saveFileDialog.FileName);
				}
			}
			catch (Exception) {
				DialogResult result = MessageBox.Show(new Form { TopMost = true }, "Failed to open save file dialog.\nTry again?", "Error:", MessageBoxButtons.YesNo);
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
		#endregion
		#region GetTextReplacementOptions
		/// <summary>
		/// Parses the command given by the user in the word document template, 
		/// and compiles it into a single TextReplacementOptions struct that can be more easily analyzed later
		/// The TextReplacementOptions struct defines exactly how the data from the CSV file will be interpretted and processed.
		/// </summary>
		public TextReplacementOptions GetTextReplacementOptions(string rawText) {
			TextReplacementOptions textReplacementOptions = new TextReplacementOptions { rawInput = rawText };
			string[] options;
			string outputType;
			string outputOption1;
			List<string> columnNames = new List<string>();
			//split the raw command into all of its various components (NOTE: the user command needs to be semicolon seperated)
			try {
				options = rawText.Split(';');
				outputType = options[0].ToLower();
				outputOption1 = options[1].ToLower();
			}
			catch (Exception) {
				throw new Exception("Error in word document input at:\n" + rawText);
			}
			//If the command is a declaration of the color pallette then parse the command as such and generate a color pallette as an array of rgb colors.
			if (outputType.Contains("colors") || outputType.Contains("colorpalette")) {
				//Parse the color pallette out of the collor pallette command and set the classwide color pallette to the new one.
				ParseAndSetColorPallette(options);
				//Set a flag to tell the rest of the program that this command was just a color pallette declaration
				textReplacementOptions.isColors = true;
				return textReplacementOptions;
			}
			//Proccess the command as an 'order' command which specifies the order in which the columns and legend entries will shop up in on the graph.
			//The position of the columns and legend entries are identical to the position of the items in this 'order' list.
			if (outputType.Contains("order")) {
				List<string> order = new List<string>();
				//Iterate through the list of items in the command and add them to order List.
				for (int i = 1; i < options.Length; i++) {
					order.Add(options[i].Trim(' ', '}'));
				}
				textReplacementOptions.order = order;
				//Set a flag that tells the rest of the program that this was just an order command. 
				//Keeps the rest of the program from attempting to process the command as a text or image replacement.
				textReplacementOptions.isOrder = true;
				return textReplacementOptions;
			}

			//Check to see whether this command asks for a bar graph, pie chart, or a simple text replacement
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
			//Check to see if this command wants the output to be a data range, mean, percentage, count, or just whatever value is unanimous in the specified columns.
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
			//If it was decided that the command asked for graph generation, then interpret the rest of the command as such.
			if (textReplacementOptions.isGraph) {
				int fontSize = 0;
				int optionsIndex = 2;
				//Check to see if the next parameter given in the command is a font size (if it is an int then it is a font size) and 
				//if it is then store it as the font size for the command.
				if (int.TryParse(options[2].Trim(' ', '}'), out fontSize)) {
					optionsIndex++;
				}
				textReplacementOptions.fontSize = fontSize;
				//All of the following parameters given are column names, so iterate through them, process them, and store them in a columnNames list.
				for (int i = optionsIndex; i < options.Length - 1; i++) {
					columnNames.Add(options[i].Trim(' ', '}'));
				}
				//The final parameter in the user-given command is always the graph title.
				textReplacementOptions.graphTitle = options[options.Length - 1].Trim(' ', '}');
			}
			//If they only provided a column name then put the column name in the name array
			else if(textReplacementOptions.isColumnValue){
				columnNames.Add(outputOption1.Trim(' ', '}'));
			}
			//If it is a simple text replacement command, then proceed to add all of the column names to the columnNames list.
			else {
				for (int i = 2; i < options.Length; i++) {
					columnNames.Add(options[i].Trim(' ', '}'));
				}
			}
			textReplacementOptions.columnNames = columnNames;
			return textReplacementOptions;
		}
		#endregion
		#region ParseAndSetColorPallette
		/// <summary>
		/// Iterates through each color entry in a given color pallette command and converts them to a color object array that can be easily used for application during graphing.
		/// </summary>
		private void ParseAndSetColorPallette(string[] options) {
			System.Drawing.Color[] colorPallete = new System.Drawing.Color[options.Length - 1];
			for (int i = 0; i < options.Length - 1; i++) {
				string[] rgb = options[i + 1].Split(',');
				//Assemble an array of 3 integers that represents an RGB color
				int[] rgbInts = new int[3];
				for (int j = 0; j < rgb.Length; j++) {
					rgbInts[j] = Convert.ToInt32(rgb[j].Trim(' ', '}'));
				}
				try {
					colorPallete[i] = System.Drawing.Color.FromArgb(rgbInts[0], rgbInts[1], rgbInts[2]);
				}
				catch (Exception e) {
					throw new Exception("Error: Invalid color in color pallette.\n" + e.Message);
				}
			}
			ColorPalette = colorPallete;
		}
		#endregion
		#region GetCommand
		/// <summary>
		/// Gets the next sequential command from the template word document and returns that command as a string.
		/// Commands are processed from the beginning of the word document to the end, so this function searches for the first occurence of a command.
		/// </summary>
		private string GetCommand(Word.Application wordApp) {
			//Initialize the Find object for finding a command.
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
				//Find the beginning of the first command in the document (denoted by three consecutive curley braces '{{{')
				sel.Find.Text = "{{{";
				if (!sel.Find.Execute()) {
					return "EOF";
				}
				Word.Range range = sel.Range;
				//Iterates through the command one character at a time until it contains the denotation for the end of the command '}}}'
				while (!sel.Text.Contains("}}}")) {
					sel.SetRange(range.Start, range.End + i);
					i++;
				}
				return sel.Text;
			}
			//Catch the case that any portion of selection fails in the document.
			catch (Exception) {
				MessageBox.Show(new Form { TopMost = true }, "Error: Failed to find text to replace in document");
				return "error";
			}
		}
		#endregion
		#region ProcessDocumentCommand
		/// <summary>
		/// Takes a command provided by the user in the word document, parses it into a TextReplacementOptions object, and executes that command based on the user-provided parameters.
		/// </summary>
		/// <param name="rawCommand">The exact text provided by the user in the document command</param>
		/// <param name="columnValueCounters">List of objects that contain information about each column</param>
		private void ProcessDocumentCommand(string rawCommand, Word.Application wordApp, List<ColumnValueCounter> columnValueCounters) {
			//Remove any extraneous braces
			string command = rawCommand.Trim('{', '}');
			TextReplacementOptions processedCommand = new TextReplacementOptions();
			//Parse the raw command to get a textReplacementOptions object that stores more usable command options.
			try {
				processedCommand = GetTextReplacementOptions(command);
			}
			catch (Exception e) {
				MessageBox.Show(new Form { TopMost = true }, "Error: " + e.Message);
				wordApp.Quit();
				Environment.Exit(1);
			}
			//if the command was a declaration of the color pallette used in the document, just remove the text and do not execute anything.
			if (processedCommand.isColors) {
				ReplaceTextWithText(rawCommand, "", wordApp);
				return;
			}
			else if (processedCommand.isOrder) {
				_graphColumnOrder = processedCommand.order;
				ReplaceTextWithText(rawCommand, "", wordApp);
				return;
			}
			//A list of the columns that we actually use for this command.
			DocumentCommandExecuter commandExecuter = new DocumentCommandExecuter(ColorPalette, _graphColumnOrder);
			//Populate a list of columnValueCounters that pertain to the columns that we are actually using for this command.
			List <ColumnValueCounter> usedColumns = GetUsedColumns(processedCommand, columnValueCounters);
			//Ensure that all columns in the list of used columns have the same number of unique row values
			NormalizeColumns(usedColumns);
			if (processedCommand.hasFailed) {
				MessageBox.Show(new Form { TopMost = true }, "Command Failed: Check template command syntax");
				wordApp.Quit();
				Environment.Exit(1);
			}
			if (usedColumns.Count == 0) {
				MessageBox.Show(new Form { TopMost = true }, "No columns found in CSV file for command:\n" + processedCommand.rawInput);
				wordApp.Quit();
				Environment.Exit(1);
			}
			//Process the command as a graph
			if (processedCommand.isGraph) {
				string replaceWith = commandExecuter.GenerateGraph(usedColumns, Application.StartupPath + "\\tempGraph.PNG", processedCommand, wordApp);
				ReplaceTextWithImage(rawCommand, replaceWith, wordApp);
			}
			//Process the command as text replacement
			else if (processedCommand.isText) {
				string replaceWith = commandExecuter.GenerateText(usedColumns, processedCommand);
				ReplaceTextWithText(rawCommand, replaceWith, wordApp);
			}
			else {
				return;
			}
		}
		#endregion
		#region ProcessDocument
		/// <summary>
		/// Searches through the document for each command, processes the commands, then prompts the user to save the changed document.
		/// </summary>
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
		#endregion
		#region GetUsedColumns
		public List<ColumnValueCounter> GetUsedColumns(TextReplacementOptions processedCommand, List<ColumnValueCounter> allColumns) {
			List<ColumnValueCounter> usedColumns = new List<ColumnValueCounter>();
			for (int i = 0; i < allColumns.Count; i++) {
				if (processedCommand.columnNames.Contains(allColumns[i].columnName.ToLower())
					|| processedCommand.columnNames.Contains(allColumns[i].columnName)
					|| processedCommand.columnNames.Contains(allColumns[i].abbreviatedRepresentation)) {
					usedColumns.Add(allColumns[i]);
				}
			}
			return usedColumns;
		}
		#endregion
		#region NormalizeColumns
		/// <summary>
		/// Ensures that each column in the list of columnValueCounters has the same 
		/// amount of unique row values as each other data column being graphed.
		/// If there is a column that does not have any occurence of one of the unique 
		/// row values that shows up in one of the other columns then it is created in that column with a count of zero.
		/// </summary>
		/// <param name="usedColumns">List of ColumnValueCounters to normalize</param>
		public void NormalizeColumns(List<ColumnValueCounter> usedColumns) {
			int uniqueRowValueCount = 0;
			//Get the column that has the most unique row values to compare the other columns against.
			//The column with the most unique row values is treated as the optimal column that we will try to match the others to.
			ColumnValueCounter maxColumn = new ColumnValueCounter();
			foreach (ColumnValueCounter column in usedColumns) {
				if (column.uniqueRowValues.Count > uniqueRowValueCount) {
					uniqueRowValueCount = column.uniqueRowValues.Count;
					maxColumn = column;
				}
			}
			if (usedColumns.Count > 1) {
				//For each used data column
				for (int i = 0; i < usedColumns.Count; i++) {
					ColumnValueCounter currentColumn = usedColumns[i];
					//If the number of unique row values of the current column is less than any of the others
					if (currentColumn.uniqueRowValues.Count < uniqueRowValueCount) {
						//scan through each row value that is in the collumn with the most unique row values
						for (int j = 0; j < maxColumn.uniqueRowValues.Count; j++) {
							//scan through each row value in the column with missing unique row values
							for (int k = 0; k < maxColumn.uniqueRowValues.Count; k++) {
								//check to see if the unique row value that we are currently looking at from the list 
								//of intended unique row values is also in the column with the list of unique row values that is missing a row value.
								if (currentColumn.uniqueRowValues[k].name == maxColumn.uniqueRowValues[j].name) {
									break;
								}
								//If we did not find the unique row value in this column then add it to the column while ensuring that it is added at the correct index in the list.
								//Making sure the index that it is at is important when we start creating chart data series using these columnValueCounters
								if (k == (currentColumn.uniqueRowValues.Count - 1)) {
									List<UniqueRowValue> tempUniqueRowValues = new List<UniqueRowValue>();
									//Populate a list of temporary uniqueRowValues with the unique row values that we have already processes/looked at 
									for (int l = 0; l < j; l++) {
										tempUniqueRowValues.Add(currentColumn.uniqueRowValues[l]);
									}
									//Add the unique row value that was excluded from this column in the correct position in the list
									tempUniqueRowValues.Add(new UniqueRowValue(maxColumn.uniqueRowValues[j].name, 0));
									//Fill the rest of the unique row values
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
		}
		#endregion
		#region Struct-TextReplacementOptions
		/// <summary>
		/// Struct that holds the color pallette, a list of column names, and 
		/// all of the other information that is relavent to deciding what to do with a user command.
		/// </summary>
		public struct TextReplacementOptions {
			/// <summary>Determines whether a command is declaring a color pallette for the document or not.
			public bool isColors;
			/// <summary>Determines whether a command is declaring the order in which values should appear in a graph.</summary>
			public bool isOrder;
			//set of descriptive booleans that determine how the data will be presented.
			public bool isGraph;
			public bool isText;
			public bool isCount;
			public bool isPercentage;
			public bool isMean;
			public bool isRange;
			public bool isColumnValue;
			/// <summary>Set when a command fails to be parsed into one of these objects </summary>
			public bool hasFailed;
			/// <summary>x-axis label font size </summary>
			public int fontSize;
			public string graphType;
			public string graphTitle;
			public string rawInput;
			public List<string> order;
			public List<string> columnNames;
			/// <summary>Stores a color pallette found in a color command. Not set if command is not a color pallette</summary>
			public System.Drawing.Color[] colorPalette;
		}
		#endregion
	}
}
