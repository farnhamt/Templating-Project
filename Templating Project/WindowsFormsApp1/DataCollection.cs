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
using Color = System.Drawing.Color;
namespace TemplatingProject {
	public class DataCollection {
		#region ClassWideVariables
		public static DataTable allData;
		public static List<string> columnHeaders;
		#endregion
		//Testing for the time being, but this generates a simple bar chart. Will need to greatly improve style to meet criteria.
		public static string generateGraph(List<ColumnValueCounter> columnValueCounters, string filename, DocumentManipulation.TextReplacementOptions options, Word.Application wordApp) {
			//Need to get the data that needs to be in the graph here.
			//allData = Form1.allData

			if (options.graphType == "bar") {
				return GenerateBarGraph(columnValueCounters, filename, options);
			}
			return filename;
			/*else if (options.graphType == "pie") {
				return GeneratePieChart(filename);
			}*/
			
		}
		#region ImportCSV
		public static DataTable ImportCSV() {
			OpenFileDialog selectFile = new OpenFileDialog();
			selectFile.Filter = "CSV Files (*.csv)|*.csv";
			if (selectFile.ShowDialog() == DialogResult.OK) {
				string CSVFilePathName = selectFile.FileName;
				DataTable dt = new DataTable();
				List<string> headers = new List<string>();
				string[] firstHeaderRow;
				string[] secondHeaderRow;
				try {
					using (StreamReader sr1 = new StreamReader(CSVFilePathName)) {
						//Get all of the column header names and place them into the data table as Columns.

						firstHeaderRow = sr1.ReadLine().Split(',');
						secondHeaderRow = sr1.ReadLine().Split(',');
						for (int i = 0; i < secondHeaderRow.Length; i++) {
							if (secondHeaderRow[i] == null || secondHeaderRow[i].Length == 0) {
								continue;
							}
							if (secondHeaderRow[i][0] == '\"') {
								secondHeaderRow[i] = secondHeaderRow[i].Split('\"')[1] /*+ ',' */ + secondHeaderRow[i + 1].Split('\"')[0];

								for (int j = (i + 2); j < secondHeaderRow.Length; j++) {
									secondHeaderRow[j - 1] = secondHeaderRow[j];
								}
								secondHeaderRow[secondHeaderRow.Length - 1] = null;
							}
						}
						for (int i = 0; i < Math.Min(secondHeaderRow.Length, firstHeaderRow.Length); i++) {
							if (secondHeaderRow[i] != "Response" && secondHeaderRow[i] != "" && secondHeaderRow[i] != null && secondHeaderRow[i] != "Open-Ended Response") {
								headers.Add(secondHeaderRow[i]);
							}
							else if (firstHeaderRow[i] != "") {
								headers.Add(firstHeaderRow[i]);
							}
						}
						int colCount = 0;
						//For each column of data add a column to the data table.
						foreach (string header in headers) {
							if (dt.Columns.Contains(header))
								dt.Columns.Add(header + "copy");
							else
								dt.Columns.Add(header);
							colCount++;
						}

					}
					using (StreamReader sr = new StreamReader(CSVFilePathName)) {

						DataRow dr = dt.NewRow();
						dr = dt.NewRow();
						string[] rows = sr.ReadLine().Split();
						rows = sr.ReadLine().Split();
						while (!sr.EndOfStream) {
							//Get a row of data
							rows = sr.ReadLine().Split(',');
							//make a data table row
							dr = dt.NewRow();
							//place all of the data in the current CSV row into our data table row one element at a time.
							for (int i = 0; i < headers.Count; i++) {
								if (rows[i] != "Response" && rows[i] != "Open-Ended Response") {
									dr[i] = rows[i];
								}
							}
							//add our row to the data table
							dt.Rows.Add(dr);
						}
					}
					columnHeaders = headers;
					return dt;

				}
				catch (System.IO.IOException) {
					MessageBoxButtons errorBoxButtons = MessageBoxButtons.OK;
					MessageBox.Show("Error: Cannot open CSV file while it is in use by another program", "File Access Error", errorBoxButtons);
					return null;
				}
			}
			else {
				return null;
			}
		}
		#endregion
		public static int countUniqueRows(DataTable allData, string columnName, string uniqueRowName) {
			string filterExpression = "[" + columnName + "]" + " = '" + uniqueRowName + "\'";
			return allData.Select(filterExpression).Length;
		}
		#region assembleColumnValueCounters
		public static List<ColumnValueCounter> assembleColumnValueCounters() {
			//columnValueCounters stores a ColumnValueCounter for every column in the data table.
			//This allows for storing the names and number of occurences of each unique data row value in relation to the column that it is a part of.
			List<ColumnValueCounter> columnValueCounters = new List<ColumnValueCounter>();
			//columnHeaders is a list of all of the column headers that we get from the importCSV function
			List<string> columnHeaders = DataCollection.columnHeaders;
			//For every column, make a new ColumnValueCounter. 
			//For each unique row value in the column, count the number of occurences of that value and store both the value and the count in the ColumnValueCounter UniqueRowValue attribute.
			
			for (int i = 0; i < allData.Columns.Count; i++) {
				ColumnValueCounter currentColumn = new ColumnValueCounter();
				currentColumn.columnName = columnHeaders[i];
				currentColumn.totalColumnValues = allData.AsDataView().ToTable(false, columnHeaders[i]).Rows.Count;
				DataRowCollection uniqueRows = allData.AsDataView().ToTable(true, columnHeaders[i]).Rows;
				//for each unique row value in this column
				UniqueRowValue unknownRowValue = new UniqueRowValue("Unknown", countUniqueRows(allData, currentColumn.columnName, ""));
				for (int j = 0; j < uniqueRows.Count; j++) {
					//Gets the name of the unique row value
					string uniqueRowName = uniqueRows[j].ItemArray[0].ToString();
					//Exclude unique row values that we do not care about (the column header and any blank rows)

					if (uniqueRowName != "" && uniqueRowName != currentColumn.columnName) {
						currentColumn.uniqueRowValues.Add(new UniqueRowValue(uniqueRowName, countUniqueRows(allData, currentColumn.columnName, uniqueRowName)));
					}
					
				}
				currentColumn.uniqueRowValues.Sort((x, y) => x.name.CompareTo(y.name));
				if (unknownRowValue != null) {
					//currentColumn.uniqueRowValues.Add(unknownRowValue);
				}
				//store that entire column
				columnValueCounters.Add(currentColumn);
			}
			return columnValueCounters;
		}
		#endregion
		#region GenerateBarGraph
		public static string GenerateBarGraph(List<ColumnValueCounter> columnValueCounters, string filename, DocumentManipulation.TextReplacementOptions options) {

			Chart chart = new Chart();
			int numSeries = options.columnNames.Count();

			List<Charting.Series> allSeries = new List<Charting.Series>();
			string[] xValues;
			if (columnValueCounters.Count > 1) {
				xValues = new string[columnValueCounters.Count];
				for (int i = 0; i < columnValueCounters.Count; i++) {
					allSeries.Add(new Charting.Series("" + i));
					xValues[i] = columnValueCounters[i].columnName;
				}
			}
			else {
				xValues = new string[columnValueCounters[0].uniqueRowValues.Count];
				allSeries.Add(new Charting.Series("0"));
				for (int i = 0; i < columnValueCounters[0].uniqueRowValues.Count; i++) {
					xValues[i] = columnValueCounters[0].uniqueRowValues[i].name;
				}
			}
			//for every relevant column
			Color[] colorPalette = new Color[] { Color.FromArgb(215,63,9), Color.FromArgb(170,157,46), Color.FromArgb(74, 119, 60) };
			for (int i = 0; i < columnValueCounters.Count; i++) {
				List<UniqueRowValue> uniqueRowValues = columnValueCounters[i].uniqueRowValues;
				float[] yValues = new float[uniqueRowValues.Count];
				//for every unique row value in the column
				for (int j = 0; j < uniqueRowValues.Count; j++) {
					int totalNumUniqueRowValues = 0;
					//Add y values based on which mode the graph is in
					if (options.isCount) {
						if(columnValueCounters.Count > 1)
							allSeries[j % allSeries.Count].Points.AddXY(xValues[i], uniqueRowValues[j].count);
						else
							allSeries[j % allSeries.Count].Points.AddXY(xValues[j], uniqueRowValues[j].count);
						//yValues[j % allSeries.Count] = uniqueRowValues[j].count;
					}
					else if (options.isPercentage) {
						//get the total number of values in the column
						for (int k = 0; k < uniqueRowValues.Count; k++) {
							totalNumUniqueRowValues += uniqueRowValues[k].count;
						}
						allSeries[j % allSeries.Count].Points.AddXY(xValues[i], Math.Round((((float)uniqueRowValues[j].count) / ((float)totalNumUniqueRowValues)) * 100, 1));
					}
				}
				//allSeries[i].Points.DataBindXY(xValues, yValues);
				allSeries[i].ChartType = Charting.SeriesChartType.RangeColumn;
				allSeries[i]["PieLabelStyle"] = "Outside";
				allSeries[i].IsValueShownAsLabel = true;
				allSeries[i].SmartLabelStyle.Enabled = true;
				allSeries[i].SmartLabelStyle.MinMovingDistance = 5;
				allSeries[i].SmartLabelStyle.AllowOutsidePlotArea = Charting.LabelOutsidePlotAreaStyle.Yes;
				allSeries[i].SmartLabelStyle.IsMarkerOverlappingAllowed = false;
				allSeries[i].SmartLabelStyle.MovingDirection = Charting.LabelAlignmentStyles.Right;
				allSeries[i].Color = colorPalette[i % (allSeries.Count)];
				//allSeries[i].SetCustomProperty("PixelPointWidth", "10");
				//allSeries[i].SetCustomProperty("PointWidth", "0.3");
				//allSeries[i].IsVisibleInLegend = true;
				//allSeries[i].LegendText = 
				/*
				Series filler = new Series();
				filler.SetCustomProperty("PointWidth", "0.3");
				filler.Enabled = true;
				
				chart.Series.Add(filler);
				*/
				chart.Series.Add(allSeries[i]);
				
			}

			Charting.ChartArea chartArea = new Charting.ChartArea("main");
			Charting.Axis yAxis = new Charting.Axis(chartArea, Charting.AxisName.Y);
			Charting.Axis xAxis = new Charting.Axis(chartArea, Charting.AxisName.X);
			chart.ChartAreas.Add(chartArea);
			chart.ChartAreas["main"].AxisX.MajorTickMark.Enabled = false;
			chart.ChartAreas["main"].AxisX.MajorGrid.Enabled = false;
			chart.ChartAreas["main"].AxisY.MajorTickMark.Enabled = false;
			chart.ChartAreas["main"].AxisY.MajorGrid.LineWidth = 1;

			chart.ChartAreas["main"].AxisY.MinorGrid.Interval = 1;
			chart.ChartAreas["main"].AxisY.MajorGrid.Interval = 1;
			chart.ChartAreas["main"].AxisY.Interval = 1;
			if (options.isPercentage) {
				chart.ChartAreas["main"].AxisY.Maximum = 100;
				chart.ChartAreas["main"].AxisY.MinorGrid.Interval = 10;
				chart.ChartAreas["main"].AxisY.MajorGrid.Interval = 10;
				chart.ChartAreas["main"].AxisY.Interval = 10;
			}  
			chart.ChartAreas["main"].AxisY.MajorGrid.LineColor = System.Drawing.Color.LightGray;
			chart.ChartAreas["main"].AxisX.LineWidth = 0;
			chart.ChartAreas["main"].AxisY.LineWidth = 0;
			//chart.ChartAreas["main"].BorderDashStyle = Charting.ChartDashStyle.Solid;
			//chart.ChartAreas["main"].BorderWidth = 1;
			chart.BorderlineDashStyle = Charting.ChartDashStyle.Solid;
			chart.BorderlineColor = System.Drawing.Color.LightGray;
			chart.BorderlineWidth = 1;
			
			Charting.Title title = new Charting.Title();
			title.Text = options.graphTitle;
			chart.Titles.Add(title);

			

			chart.Width = 600;
			chart.Height = 300;
			//chart.PaletteCustomColors = System.Drawing.Color.FromArgb();
			try {
				chart.SaveImage(filename, Charting.ChartImageFormat.Png);
			}
			catch (Exception e) {

			}
			return filename;
		}
		#endregion
		#region generateText
		public static string generateText(List<ColumnValueCounter> usedColumns, string rawCommand, DocumentManipulation.TextReplacementOptions processedCommand, Word.Application wordApp) {
			string assembledText = "";
			int unknownCount;
			if (processedCommand.isColumnValue) {
				return usedColumns[0].uniqueRowValues[0].name;
			}
			else if (processedCommand.isCount) {
				foreach (ColumnValueCounter column in usedColumns) {
					unknownCount = column.totalColumnValues;
					foreach (UniqueRowValue row in column.uniqueRowValues) {
						if (column.uniqueRowValues.Count > 1) {
							assembledText += row.name;
							assembledText += ": ";
							assembledText += row.count;
							assembledText += ", ";
						}
						else {
							assembledText += row.count;
						}
						unknownCount -= row.count;
					}
					if (unknownCount != 0) {
						assembledText += "Unknown: ";
						assembledText += unknownCount;
					}
				}
			}
			else if (processedCommand.isRange) {
				int lowest = int.MaxValue;
				int highest = int.MinValue;
				int current;
				foreach (ColumnValueCounter column in usedColumns) {
					foreach (UniqueRowValue row in column.uniqueRowValues) {
						if ((current = wordToInt(row.name)) != -1) {
							if (current < lowest)
								lowest = current;
							if (current > highest)
								highest = current;
						}
					}
				}
				assembledText += lowest;
				assembledText += " - ";
				assembledText += highest;
			}
			else if (processedCommand.isMean) {
				int total = 0;
				int current = 0;
				int uniqueRowValueCount = 0;
				foreach (ColumnValueCounter column in usedColumns) {
					foreach (UniqueRowValue row in column.uniqueRowValues) {
						if ((current = wordToInt(row.name)) != -1) {
							total += (current * row.count);
						}
						uniqueRowValueCount += row.count;
					}
				}
				assembledText = Math.Round((float)total / (float)uniqueRowValueCount, 2).ToString();
			}
			return assembledText;
		}
		#endregion
		public static int wordToInt(string word) {
			word = word.ToLower();
			string[] numbers = {
			"zero", "one", "two", "three", "four", "five", "six", "seven", "eight",
			"nine", "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen",
			"sixteen", "seventeen", "eighteen", "nineteen", "twenty"};
			if (!(numbers.Contains(word))) {
				return -1;
			}
			switch (word) {
				case "zero":
					return 0;
				case "one":
					return 1;
				case "two":
					return 2;
				case "three":
					return 3;
				case "four":
					return 4;
				case "five":
					return 5;
				case "six":
					return 6;
				case "seven":
					return 7;
				case "eight":
					return 8;
				case "nine":
					return 9;
				case "ten":
					return 10;
				case "eleven":
					return 11;
				case "twelve":
					return 12;
				case "thirteen":
					return 13;
				case "fourteen":
					return 14;
				case "fifteen":
					return 15;
				case "sixteen":
					return 16;
				case "seventeen":
					return 17;
				case "eighteen":
					return 18;
				case "nineteen":
					return 19;
				case "twenty":
					return 20;
				default:
					return -1;
			}
		}

	}


	public class UniqueRowValue {
		public string name;
		public int count;
		public UniqueRowValue(string name, int count) {
			this.name = name;
			this.count = count;
		}
	}
	public class ColumnValueCounter {
		public List<UniqueRowValue> uniqueRowValues;
		public string columnName;
		public int totalColumnValues;
		public ColumnValueCounter() => uniqueRowValues = new List<UniqueRowValue>();
	}
}
