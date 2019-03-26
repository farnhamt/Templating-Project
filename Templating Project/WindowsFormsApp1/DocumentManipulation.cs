using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data;
using Word = Microsoft.Office.Interop.Word;
using Charting = System.Windows.Forms.DataVisualization.Charting;
using Chart = System.Windows.Forms.DataVisualization.Charting.Chart;
using Series = System.Windows.Forms.DataVisualization.Charting.Series;

namespace TemplatingProject {
	public class DocumentManipulation {
		public static Word.Application openDocument(string filepath) {
			Word.Application wordApp = new Word.Application();
			try {
				wordApp.Application.Documents.Open(filepath, false, false);
			}
			catch (Exception error) {
				MessageBox.Show("Failed to open document!\nError: " + error.Message);
				return null;
			}
			return wordApp;
		}
		
		public static void replaceTextWithText(string textToReplace, string replacementText, Word.Application wordApp) {
			object replaceAll = Word.WdReplace.wdReplaceAll;
			Word.Find findObject = wordApp.Selection.Find;
			findObject.Execute(textToReplace, true, true, false, false, false, false, Word.WdFindWrap.wdFindAsk, false, replacementText, replaceAll, false, false, false, false);
		}
		public static void replaceTextWithImage(string rawCommand, TextReplacementOptions options, string imagePath, Word.Application wordApp) {
			Word.Selection sel = wordApp.Selection;
			string keyword = rawCommand;
			sel.Find.Execute(keyword, Replace: Word.WdReplace.wdReplaceOne);
			sel.Range.Select();
			var imagePath1 = Path.GetFullPath(string.Format(imagePath, keyword));
			sel.InlineShapes.AddPicture(imagePath, false, true);
		}
		public static void saveDocumentAs(string filename, Word.Application wordApp) {
			try {
				wordApp.Application.ActiveDocument.SaveAs2(@"C:\VSTesting\" + filename + ".docx");
				wordApp.Application.ActiveDocument.SaveAs(@"C:\VSTesting\" + filename + ".pdf", Word.WdSaveFormat.wdFormatPDF);
			}
			catch (Exception e) {
				saveDocumentAs(filename, wordApp);
				wordApp.Quit();
				return;
			}
			wordApp.Quit();
		}

		public static TextReplacementOptions GetTextReplacementOptions(string rawText) {
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
				for (int i = 2; i < options.Length - 1; i++) {
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

		public struct TextReplacementOptions {
			public bool isGraph;
			public bool isText;
			public bool isCount;
			public bool isPercentage;
			public bool isMean;
			public bool isRange;
			public bool isColumnValue;

			public bool hasFailed;

			public string graphType;
			public string graphTitle;
			public string rawInput;
			public List<string> columnNames;

		}
		public static string GetTextToReplace(Word.Application wordApp) {
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

		public static void ProcessDocumentCommand(string rawCommand, Word.Application wordApp, List<ColumnValueCounter> columnValueCounters) {
			string command = rawCommand.Trim('{', '}');
			TextReplacementOptions processedCommand = GetTextReplacementOptions(command);
			List<ColumnValueCounter> usedColumns = new List<ColumnValueCounter>();
			for (int i = 0; i < columnValueCounters.Count; i++) {
				if (processedCommand.columnNames.Contains(columnValueCounters[i].columnName.ToLower()) || processedCommand.columnNames.Contains(columnValueCounters[i].columnName)) {
					usedColumns.Add(columnValueCounters[i]);
				}
			}
			if (processedCommand.hasFailed) {
				MessageBox.Show("Command Failed: Check template command syntax");
				wordApp.Quit();
			}
			if (processedCommand.isGraph) {
				string replaceWith = DataCollection.generateGraph(usedColumns, @"C:\VSTesting\tempGraph.PNG", processedCommand, wordApp);
				replaceTextWithImage(rawCommand, processedCommand, replaceWith, wordApp);
			}
			else if (processedCommand.isText) {
				string replaceWith = DataCollection.generateText(usedColumns, rawCommand, processedCommand, wordApp);
				replaceTextWithText(rawCommand, replaceWith, wordApp);
			}
			else {
				return;
			}
		}
	}

}
