using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Charting = System.Windows.Forms.DataVisualization.Charting;
using Chart = System.Windows.Forms.DataVisualization.Charting.Chart;
using Series = System.Windows.Forms.DataVisualization.Charting.Series;
using Color = System.Drawing.Color;
using System.Windows.Forms.DataVisualization.Charting;

namespace TemplatingProject {
	class DocumentCommandExecuter {
		private Color[] _colorPallette;
		#region Constructor
		public DocumentCommandExecuter(Color[] colorPallette) {
			_colorPallette = colorPallette;
		}
		#endregion
		#region GenerateGraph
		/// <summary>
		/// Generates a graph based on the columnValueCounter list input and returns the filename of a graph in PNG format.
		/// </summary>
		/// <param name="columnValueCounters"></param>
		/// <param name="filename">File name to save graph as</param>
		/// <param name="options">Object that defines the parameters of the graph</param>
		public string GenerateGraph(List<ColumnValueCounter> columnValueCounters, string filename, DocumentManipulation.TextReplacementOptions options) {
			if (options.graphType == "bar") {
				return GenerateBarGraph(columnValueCounters, filename, options);
			}
			else if (options.graphType == "pie") {
				return GeneratePieChart(columnValueCounters, filename, options);
			}
			return filename;
		}
		#endregion
		#region GenerateBarGraph
		/// <summary>
		/// Creates a bar graph based on columnValueCounter input and saves the graph as an image file under the given filename.
		/// </summary>
		private string GenerateBarGraph(List<ColumnValueCounter> columnValueCounters, string filename, DocumentManipulation.TextReplacementOptions options) {

			Chart chart = new Chart();
			List<Series> allSeries = new List<Series>();

			if (columnValueCounters.Count == 0) {
				MessageBox.Show(new Form { TopMost = true }, "Error: No data imported. Invalid CSV data format");
				Environment.Exit(1);
			}
			//Get chart series X values from data
			string[] xValues = InitializeChartSeriesXValues(allSeries, columnValueCounters, chart);

			for (int i = 0; i < columnValueCounters.Count; i++) {
				List<UniqueRowValue> uniqueRowValues = columnValueCounters[i].uniqueRowValues;
				float[] yValues = new float[uniqueRowValues.Count];

				//for every unique row value in the column
				for (int j = 0; j < uniqueRowValues.Count; j++) {
					int k;
					int totalNumUniqueRowValues = 0;
					//Add y values based on which mode the graph is in
					if (options.isCount) {
						//If there is more than one relevant column then the xValues are the column names, otherwise the xValues are the uniqueRowValue names.
						if (columnValueCounters.Count > 1)
							allSeries[j % allSeries.Count].Points.AddXY(xValues[i], uniqueRowValues[j].count);
						else
							allSeries[j % allSeries.Count].Points.AddXY(xValues[j], uniqueRowValues[j].count);
					}
					else if (options.isPercentage) {
						//get the total number of values in the column
						for (k = 0; k < uniqueRowValues.Count; k++) {
							totalNumUniqueRowValues += uniqueRowValues[k].count;
						}
						allSeries[j % allSeries.Count].Points.AddXY(xValues[i], Math.Round((((float)uniqueRowValues[j].count) / ((float)totalNumUniqueRowValues)) * 100, 1));
					}
				}
			}
			if (columnValueCounters.Count <= 1) {
				//Add another column that represents the number of "Unknown" responses (rows that were blank in that column).
				allSeries[0].Points.AddXY("Unknown", columnValueCounters[0].unknownCount);
				//Apply all of the customized settings for the series and add it to the chart.
				FinalizeBarChartSeries(chart, allSeries[0], columnValueCounters, _colorPallette[0]);
			}
			else {
				for (int j = 0; j < allSeries.Count; j++) {
					if (j > 0) {
						//To create space between different series on the chart, insert a filler series that acts as a spacer.
						CreateFillerChartSeries(chart, allSeries[j].Name, xValues.ToList());
					}
					//Apply all of the customized settings for the series and then add the series to the chart.
					FinalizeBarChartSeries(chart, allSeries[j], columnValueCounters, _colorPallette[j]);
				}
				CreateFillerChartSeries(chart, "end", xValues.ToList());
				CreateFillerChartSeries(chart, "end1", xValues.ToList());
			}
			//Apply all of the customized settings for the chart itself.
			ApplyCustomChartOptions(chart, options, columnValueCounters);

			try {
				chart.SaveImage(filename, ChartImageFormat.Png);
			}
			catch (Exception) {
				MessageBox.Show(new Form { TopMost = true }, "Error: Failed to save graph image");
			}
			return filename;
		}
		#endregion
		#region GeneratePieChart
		private string GeneratePieChart(List<ColumnValueCounter> columnValueCounters, string filename, DocumentManipulation.TextReplacementOptions options) {

			Chart chart = new Chart();
			List<Series> allSeries = new List<Series>();
			
			if (columnValueCounters.Count == 0) {
				MessageBox.Show(new Form { TopMost = true }, "Error: No data imported. Invalid CSV data format");
				Environment.Exit(1);
			}
			string[] xValues = InitializeChartSeriesXValues(allSeries, columnValueCounters, chart);

			int totalNumUniqueRowValues = 0;
			for (int i = 0; i < columnValueCounters.Count; i++) {
				List<UniqueRowValue> uniqueRowValues = columnValueCounters[i].uniqueRowValues;
				float[] yValues = new float[uniqueRowValues.Count];

				//for every unique row value in the column
				for (int j = 0; j < uniqueRowValues.Count; j++) {
					int k;
					//Add y values based on which mode the graph is in
					if (options.isCount) {
						//If there is more than one relevant column then the xValues are the column names, otherwise the xValues are the uniqueRowValue names.
						if (columnValueCounters.Count > 1)
							allSeries[j % allSeries.Count].Points.AddXY(xValues[i], uniqueRowValues[j].count);
						else
							allSeries[j % allSeries.Count].Points.AddXY(xValues[j], uniqueRowValues[j].count);
					}
					else if (options.isPercentage) {
						totalNumUniqueRowValues = 0;
						//get the total number of values in the column
						for (k = 0; k < uniqueRowValues.Count; k++) {
							totalNumUniqueRowValues += uniqueRowValues[k].count;
						}
						if (columnValueCounters.Count > 1)
							allSeries[j % allSeries.Count].Points.AddXY(xValues[i], Math.Round((((float)uniqueRowValues[j].count) / ((float)totalNumUniqueRowValues + columnValueCounters[0].unknownCount)) * 100, 1));
						else
							allSeries[j % allSeries.Count].Points.AddXY(xValues[j], Math.Round((((float)uniqueRowValues[j].count) / ((float)totalNumUniqueRowValues + columnValueCounters[0].unknownCount)) * 100, 1));

					}
				}
			}
			if (columnValueCounters.Count <= 1) {
				//Add another column that represents the number of "Unknown" responses (rows that were blank in that column).
				if (options.isCount) {
					allSeries[0].Points.AddXY("Unknown", columnValueCounters[0].unknownCount);
				}
				else if (options.isPercentage) {
					allSeries[0].Points.AddXY("Unknown", Math.Round((((float)columnValueCounters[0].unknownCount) / ((float)totalNumUniqueRowValues + columnValueCounters[0].unknownCount)) * 100, 1));
				}
				//Apply all of the customized settings for the series and add it to the chart.
				if (options.graphType == "bar")
					FinalizeBarChartSeries(chart, allSeries[0], columnValueCounters, _colorPallette[0]);
				else if (options.graphType == "pie") {
					FinalizePieChartSeries(chart, allSeries[0], columnValueCounters);
				}
			}
			else {
				for (int j = 0; j < allSeries.Count; j++) {
					if (j > 0) {
						//To create space between different series on the chart, insert a filler series that acts as a spacer.
						CreateFillerChartSeries(chart, allSeries[j].Name, xValues.ToList());
					}
					//Apply all of the customized settings for the series and then add the series to the chart.
					FinalizeBarChartSeries(chart, allSeries[j], columnValueCounters, _colorPallette[j]);
				}
			}
			//Apply all of the customized settings for the chart itself.
			ApplyCustomChartOptions(chart, options, columnValueCounters);

			try {
				chart.SaveImage(filename, ChartImageFormat.Png);
			}
			catch (Exception) {
				MessageBox.Show(new Form { TopMost = true }, "Error: Failed to create graph image");
			}
			return filename;
		}

		
		#endregion
		#region Finalize Bar and Pie Chart Series
		/// <summary>
		/// Configures all style settings for an individual bar chart data series.
		/// </summary>
		private void FinalizeBarChartSeries(Chart chart, Series series, List<ColumnValueCounter> columnValueCounters, Color color) {
			series.ChartType = SeriesChartType.Column;
			series["PieLabelStyle"] = "Outside";
			series.Color = color;
			series.IsVisibleInLegend = true;
			//Set series label style
			series.CustomProperties = "BarLabelStyle = Top";
			series.CustomProperties = "LabelStyle = Top";
			series.Font = new System.Drawing.Font("Calibri", 16);
			series.IsValueShownAsLabel = true;
			series.SmartLabelStyle.Enabled = false;
			series.SmartLabelStyle.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.Partial;
			//When we are plotting more than one column at a time, there will be filler series' in between each data series.
			//Increase the width of the columns to account for the decreased width that gets automatically applied to them because of this.
			if (columnValueCounters.Count > 1) {
				//Set width for multiple series case
				series.SetCustomProperty("PointWidth", "1");
			}
			else {
				//Set width for single series case
				series.SetCustomProperty("PointWidth", "0.3");
			}
			chart.Series.Add(series);
		}
		/// <summary>
		/// Configures all style settings for an individual pie chart data series.
		/// </summary>
		private void FinalizePieChartSeries(Chart chart, Series series, List<ColumnValueCounter> columnValueCounters) {
			series.ChartType = SeriesChartType.Pie;
			//Sort points by largest y value (percentage or count) to avoid the pie chart slices appearing in random order
			//Simple bubble sort (gaurunteed small set size)
			for (int i = 0; i < series.Points.Count; i++) {
				DataPoint temp = new DataPoint();
				for (int j = i; j < series.Points.Count; j++) {
					if (series.Points[j].YValues[0] > series.Points[i].YValues[0]) {
						temp = series.Points[i];
						series.Points[i] = series.Points[j];
						series.Points[j] = temp;
					}
				}
				series.Points[i].IsVisibleInLegend = true;
			}
			//Set colors for each point in the series basedon the provided color pallette colors
			//First check to see if there are enough colors for all of the data points. If there are not then it keeps recycling them.
			if (series.Points.Count > _colorPallette.Length) {
				MessageBox.Show(new Form { TopMost = true }, "Not enough colors in color palette to satisfy all data points");
				for (int i = 0; i < series.Points.Count; i++) {
					series.Points[i].Color = _colorPallette[i % _colorPallette.Length];
				}
			}
			//Otherwise just 1:1 map colors to points in order that they appear in the list.
			else {
				for (int i = 0; i < series.Points.Count; i++) {
					series.Points[i].Color = _colorPallette[i];
				}
			}
			series.IsVisibleInLegend = true;
			//Configure label style
			series.IsValueShownAsLabel = true;
			series.SmartLabelStyle.Enabled = true;
			series.CustomProperties = "BarLabelStyle = Top";
			series.CustomProperties = "LabelStyle = Top";
			series.Font = new System.Drawing.Font("Calibri", 16);

			chart.Series.Add(series);
		}
		#endregion
		#region CreateFillerChartSeries
		/// <summary>
		/// Creates a filler series for the given chart to provide padding between bars in a bar graph.
		/// Preconditions: Chart must be bar graph and have more than one desired series of data (columnValueCounters.Count > 1).
		/// </summary>
		private void CreateFillerChartSeries(Chart chart, string fillerName, List<string> xValues) {
			Series filler = new Series("filler" + fillerName);
			//Need to set 0 for the y-value corresponding to each xValue to make sure this series does not actually show up in the chart.
			foreach (string xValue in xValues) {
				filler.Points.AddXY(xValue, 0);
			}
			filler.Enabled = true;
			filler.IsVisibleInLegend = false;
			filler.SetCustomProperty("PointWidth", "0.1");
			chart.Series.Add(filler);
		}
		#endregion
		#region ApplyCustomChartOptions
		/// <summary>
		/// Applies stylistic configuration to the given chart based on the TextReplacementOptions that are passed in.
		/// </summary>
		private void ApplyCustomChartOptions(Chart chart, DocumentManipulation.TextReplacementOptions options, List<ColumnValueCounter> columnValueCounters) {
			//Initialize chart area and axis
			ChartArea chartArea = new ChartArea("main");
			Axis yAxis = new Axis(chartArea, AxisName.Y);
			Axis xAxis = new Axis(chartArea, AxisName.X);
			chart.ChartAreas.Add(chartArea);
			//Configure Y-Axis Style
			chart.ChartAreas["main"].AxisY.MajorTickMark.Enabled = false;
			chart.ChartAreas["main"].AxisY.MajorGrid.LineWidth = 1;
			chart.ChartAreas["main"].AxisY.MinorGrid.Interval = 1;
			chart.ChartAreas["main"].AxisY.MajorGrid.LineColor = Color.Gray;
			chart.ChartAreas["main"].AxisY.MajorGrid.Interval = 1;
			chart.ChartAreas["main"].AxisY.Interval = 1;
			if (options.isPercentage) {
				chart.ChartAreas["main"].AxisY.Maximum = 100;
				chart.ChartAreas["main"].AxisY.MinorGrid.Interval = 10;
				chart.ChartAreas["main"].AxisY.MajorGrid.Interval = 10;
				chart.ChartAreas["main"].AxisY.Interval = 10;
			}
			chart.ChartAreas["main"].AxisY.MajorGrid.LineColor = Color.LightGray;
			chart.ChartAreas["main"].AxisY.LineWidth = 0;
			//Configure Y-Axis label style
			chart.ChartAreas["main"].AxisY.LabelAutoFitMinFontSize = 16;
			chart.ChartAreas["main"].AxisY.LabelAutoFitMaxFontSize = 16;
			//Configure X-Axis style
			chart.ChartAreas["main"].AxisX.MajorTickMark.Enabled = false;
			chart.ChartAreas["main"].AxisX.MajorGrid.Enabled = false;
			chart.ChartAreas["main"].AxisX.LineWidth = 0;
			//Configure x-axis label style
			chart.ChartAreas["main"].AxisX.IsLabelAutoFit = false;
			chart.ChartAreas["main"].AxisX.LabelStyle.Font = new System.Drawing.Font("Calibri", 16);
			if (options.fontSize <= 0) {
				chart.ChartAreas["main"].AxisX.IsLabelAutoFit = true;
			}
			else {
				chart.ChartAreas["main"].AxisX.LabelStyle.Font = new System.Drawing.Font("Calibri", options.fontSize);
			}
			chart.ChartAreas["main"].AxisX.LabelAutoFitMaxFontSize = 16;
			chart.ChartAreas["main"].AxisX.LabelAutoFitMinFontSize = 16;
			chart.ChartAreas["main"].AxisX.LabelStyle.ForeColor = Color.Black;
			//Configure chart border style
			chart.BorderlineDashStyle = ChartDashStyle.Solid;
			chart.BorderlineColor = Color.LightGray;
			chart.BorderlineWidth = 1;
			//Ensure chart antialiasing is off
			chart.AntiAliasing = AntiAliasingStyles.None;
			//Configure graph legend
			Legend legend = new Legend {
				Font = new System.Drawing.Font("Calibri", 16),
				IsTextAutoFit = false,
				Alignment = System.Drawing.StringAlignment.Center,
				LegendStyle = LegendStyle.Row
			};
			if (columnValueCounters.Count > 1) {
				legend.Docking = Docking.Bottom;
				chart.Legends.Add(legend);
			}
			if (options.graphType == "pie") {
				legend.LegendStyle = LegendStyle.Table;
				legend.Docking = Docking.Right;
				legend.Font = new System.Drawing.Font("Calibri", 16);
				legend.IsEquallySpacedItems = true;
				legend.TableStyle = LegendTableStyle.Wide;
				//legend.MaximumAutoSize = 100;
				//legend.AutoFitMinFontSize = 32;
				//legend.IsTextAutoFit = true;
				chart.Legends.Add(legend);
			}
			//Configure graph title
			Title title = new Title {
				Text = options.graphTitle,
				Font = new System.Drawing.Font("Calibri", 24, System.Drawing.FontStyle.Italic),
				ForeColor = Color.Gray
			};
			chart.Titles.Add(title);
			//Configure pie chart specific chart settings
			if (options.graphType == "pie") {
				//Set the font size of the data labels that reside inside of the pie chart
				chart.Series[0].Font = new System.Drawing.Font("Calibri", 16);
				chart.BorderlineWidth = 0;
				chart.Width = 720;
				chart.Height = 560;
			}
			//Configure bar chart with multiple data columns specific settings
			else if (columnValueCounters.Count > 1) {
				chart.Width = 1200;
				chart.Height = 600;
			}
			//Configure bar chart that uses a single data column specific settings
			else {
				chart.Width = 1000;
				chart.Height = 400;
			}
		}
		#endregion
		#region GenerateText
		public string GenerateText(List<ColumnValueCounter> usedColumns, string rawCommand, DocumentManipulation.TextReplacementOptions processedCommand, Word.Application wordApp) {
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
						if ((current = WordToInt(row.name)) != -1) {
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
						if ((current = WordToInt(row.name)) != -1) {
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
		#region InitializeChartSeriesXValues
		/// <summary> Initializes the xvalues of the chart based on whether the graph has multiple series or a single series. </summary>
		private string[] InitializeChartSeriesXValues(List<Series> allSeries, List<ColumnValueCounter> columnValueCounters, Chart chart) {
			string[] xValues;
			
			//If this graph is being created using multiple columns then place the xValues into different series', otherwise, place them all into one series.
			
			//The x-values in this case are the names of the columns that are used.
			if (columnValueCounters.Count > 1) {
				xValues = new string[columnValueCounters.Count];
				//Create a new series for each data column
				for (int i = 0; i < columnValueCounters[0].uniqueRowValues.Count; i++) {
					allSeries.Add(new Series(columnValueCounters[0].uniqueRowValues[i].name));
				}
				//Assign xValues as the column name of each column
				for (int i = 0; i < columnValueCounters.Count; i++) {
					xValues[i] = columnValueCounters[i].columnName;
				}
				//Create filler series' to create more space in between each set of columns associated with x-values in the chart
				CreateFillerChartSeries(chart, "beginning", xValues.ToList());
				CreateFillerChartSeries(chart, "beginning1", xValues.ToList());
			}
			//The x-values are the names of the unique row values in the single column.
			else {
				//If there are any blank/unknown rows in the column, then create another x-value that represents the unknown values.
				if (columnValueCounters[0].unknownCount > 0) {
					xValues = new string[columnValueCounters[0].uniqueRowValues.Count + 1];
					xValues[columnValueCounters[0].uniqueRowValues.Count] = "Unknown";
				}
				else {
					xValues = new string[columnValueCounters[0].uniqueRowValues.Count];
				}
				//Create the single series for this chart.
				allSeries.Add(new Series("0"));
				for (int i = 0; i < columnValueCounters[0].uniqueRowValues.Count; i++) {
					xValues[i] = columnValueCounters[0].uniqueRowValues[i].name;
				}
			}
			return xValues;
		}
		#endregion
		#region WordToInt
		/// <summary>
		/// Takes in a number from 0 to 20 as an english word string and returns its corresponding integer value.
		/// Used to process data that comes in as this word format as numerical values.
		/// </summary>
		private int WordToInt(string word) {
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
		#endregion
	}
}
