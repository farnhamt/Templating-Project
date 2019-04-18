using System.Collections.Generic;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;


namespace TemplatingProject {
	public partial class Main : Form {
		private DataCollection _dataCollector = new DataCollection();
		private DocumentManipulation _documentManipulator = new DocumentManipulation();
		public Main()
        {
			//Prompt user to select the word document template they would like to use.
			Word.Application wordApp = OpenTemplate();

			//Prompt user to select CSV file and import the data from it.
			if (!_dataCollector.ImportCSV()) {
				wordApp?.Quit();
				System.Environment.Exit(1);
			}

			List<ColumnValueCounter> columnValueCounters = _dataCollector.assembleColumnValueCounters();
			_documentManipulator.ProcessDocument(wordApp, columnValueCounters);
			
			MessageBox.Show("done");
			System.Environment.Exit(0);
		}
		#region OpenTemplate
		/// <summary>
		/// Prompts the user to select the word document that they want to use as a template and then creates a new Word.Application by opening that file.
		/// </summary>
		private Word.Application OpenTemplate() {
			OpenFileDialog selectFile = new OpenFileDialog();
			selectFile.Filter = "Word 2007 Documents (*.docx)|*.docx| Word 97-2003 Documents (*.doc)|*.doc";
			selectFile.AutoUpgradeEnabled = false;
			if (selectFile.ShowDialog() == DialogResult.OK) {
				return _documentManipulator.OpenDocument(selectFile.FileName);
			}
			else {
				MessageBox.Show("Error: Failed to open word document");
				System.Environment.Exit(1);
				return null;
			}
		}
		#endregion
	}

}