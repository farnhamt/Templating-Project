using System;
using System.Collections.Generic;
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


namespace TemplatingProject
{
	public partial class Form1 : Form {
		public Form1()
        {
            InitializeComponent();
        }
        private void TemplateButtonClick(object sender, EventArgs e)
        {
            string senderString = sender.ToString().Split(':')[1].Substring(1);
            //Prompt user to select CSV file and import the data from it.
            DataCollection.allData = DataCollection.ImportCSV();
            //If there was no data found then error out.
            if (DataCollection.allData == null){
                MessageBoxButtons errorBoxButtons = MessageBoxButtons.OK;
                MessageBox.Show("Error: No data imported from CSV file", "Error detected in input", errorBoxButtons);
                return;
            }
			Word.Application wordApp = DocumentManipulation.openDocument(@"C:\VSTesting\Civic Engagement.docx");
			List <ColumnValueCounter> columnValueCounters = DataCollection.assembleColumnValueCounters();
			//NOTE TO SELF: need to calculate text replacement options in this class using the column value counters that we have. Then generate the graphs. THEN pass them to document manipulation to do the actual text replacement.
			/*for (int i = 0; i < columnValueCounters.Count; i++) {
				if (columnValueCounters[i].uniqueRowValues.Count > 1) {
					DocumentManipulation.replaceText(@"C:\VSTesting\tempGraph" + i + ".PNG", columnValueCounters[i].columnName, wordApp, columnValueCounters);
				}
				else if(columnValueCounters[i].uniqueRowValues.Count == 1){
					DocumentManipulation.replaceText(columnValueCounters[i].uniqueRowValues[0].name, columnValueCounters[i].columnName, wordApp, false);
				}
			}*/
			string documentCode = "";
			while (true) {
				
				documentCode = DocumentManipulation.GetTextToReplace(wordApp);
				if (documentCode == "EOF") {
					break;
				}
				else if (documentCode == "error") {
					wordApp.Quit();
					return;
				}
				else {
					DocumentManipulation.ProcessDocumentCommand(documentCode, wordApp, columnValueCounters);
				}

			}
			DocumentManipulation.saveDocumentAs("helloWorld", wordApp);
			
            MessageBox.Show("done");
        }

    }

}