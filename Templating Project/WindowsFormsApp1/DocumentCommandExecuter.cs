using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Chart = System.Windows.Forms.DataVisualization.Charting.Chart;
using Series = System.Windows.Forms.DataVisualization.Charting.Series;
using Color = System.Drawing.Color;
using System.Windows.Forms.DataVisualization.Charting;

namespace TemplatingProject {
	/// <summary>
	/// Generates the text and graph images that will be placed into the output document.
	/// </summary>
	class DocumentCommandExecuter {
		#region Class Wide Variables
		/// <summary>
		/// The set of colors to apply to the generated graphs.
		/// </summary>
		private Color[] _colorPallette;
		/// <summary>
		/// An array of strings that correspond to data row values in the order that the user wants them to appear in the graphs.
		/// </summary>
		private List<string> _itemOrder;
		#endregion
		#region Constructor
		public DocumentCommandExecuter(Color[] colorPallette, List<string> itemOrder = null) {
			_colorPallette = colorPallette;
			_itemOrder = itemOrder;
		}
		#endregion
		#region GenerateGraph
		/// <summary>
		/// Generates a graph based on the columnValueCounter list input and returns the filename of a graph in PNG format.
		/// </summary>
		/// <param name="columnValueCounters"></param>
		/// <param name="filename">File name to save graph as</param>
		/// <param name="options">Object that defines the parameters of the graph</param>
		public string GenerateGraph(List<ColumnValueCounter> columnValueCounters, string filename, DocumentManipulation.TextReplacementOptions options, Word.Application wordApp) {
			try {
				if (options.graphType == "bar") {
					return GenerateBarGraph(columnValueCounters, filename, options);
				}
				else if (options.graphType == "pie") {
					return GeneratePieChart(columnValueCounters, filename, options);
				}
			}
			catch (Exception e) {
				MessageBox.Show(new Form { TopMost = true }, "Error: " + e.Message);
				wordApp.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
				Environment.Exit(1);
			}
			//returns the file name as confirmation that the graph image is saved in the correct file location.
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
				throw new Exception("No data imported. Invalid CSV data format");
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
						if (columnValueCounters.Count > 1) {
							allSeries[j % allSeries.Count].Points.AddXY(xValues[i], uniqueRowValues[j].count);
							allSeries[j % allSeries.Count].Name = uniqueRowValues[j].name;
						}
						else {
							allSeries[j % allSeries.Count].Points.AddXY(xValues[j], uniqueRowValues[j].count);
						}
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
				//Reorder the data values so they show up in the appropriate as specified by the user.
				ApplyItemOrder(allSeries[0]);
				//Apply all of the customized settings for the series and add it to the chart.
				FinalizeBarChartSeries(chart, allSeries[0], columnValueCounters, _colorPallette[0], options);
			}
			else {
				//Place the series in the correct order as specified by the user in an 'order' command.
				ApplyItemOrder(allSeries);
				
				//for each series, create spacing between the bars in the chart, apply a color from the color pallette, and apply custom series settings for the bar chart.
				for (int j = 0; j < allSeries.Count; j++) {
					if (j > 0) {
						//To create space between different series on the chart, insert a filler series that acts as a spacer.
						CreateFillerChartSeries(chart, allSeries[j].Name, xValues.ToList());
					}
					if (j >= _colorPallette.Length) {
						MessageBox.Show(new Form { TopMost = true }, "Warning: Not enough colors specified in color pallette to accomodate data input.");
					}
					Color seriesColor = _colorPallette[j % _colorPallette.Length];
					//Apply all of the customized settings for the series and then add the series to the chart.
					FinalizeBarChartSeries(chart, allSeries[j], columnValueCounters, seriesColor, options);
				}
				//Add more space between each set of bar graph columns
				CreateFillerChartSeries(chart, "end", xValues.ToList());
				CreateFillerChartSeries(chart, "end1", xValues.ToList());
			}
			//Apply all of the customized settings for the chart itself.
			ApplyCustomChartOptions(chart, options, columnValueCounters);

			try {
				chart.SaveImage(filename, ChartImageFormat.Png);
			}
			catch (Exception) {
				throw new Exception("Failed to save graph image");
			}
			return filename;
		}
		#endregion
		#region GeneratePieChart
		private string GeneratePieChart(List<ColumnValueCounter> columnValueCounters, string filename, DocumentManipulation.TextReplacementOptions options) {

			Chart chart = new Chart();
			//Create a new series list to pass into InitializeChartSeriesXValues. Will be converted to a single series later.
			List<Series> allSeries = new List<Series>();

			if (columnValueCounters.Count == 0) {
				throw new Exception("No data imported. Invalid CSV data format");
			}
			else if (columnValueCounters.Count > 1) {
				throw new Exception("Invalid Pie Chart Schema");
			}
			//Get the xValues that will be represented in the pie chart. These are the row values within the column that is being made into a pie chart.
			string[] xValues = InitializeChartSeriesXValues(allSeries, columnValueCounters, chart);
			//Convert back to a single series
			Series pieSeries = allSeries[0];
			int totalNumUniqueRowValues = 0;
			for (int i = 0; i < columnValueCounters.Count; i++) {
				List<UniqueRowValue> uniqueRowValues = columnValueCounters[i].uniqueRowValues;
				float[] yValues = new float[uniqueRowValues.Count];

				//for every unique row value in the column
				for (int j = 0; j < uniqueRowValues.Count; j++) {
					int k;
					//Add y values based on which mode the graph is in
					if (options.isCount) {
						//Add a point to the pie chart corresponding to the number of this specific unique row value occurences.
						pieSeries.Points.AddXY(xValues[j], uniqueRowValues[j].count);
					}
					else if (options.isPercentage) {
						totalNumUniqueRowValues = 0;
						//get the total number of values in the column to use as a divisor
						for (k = 0; k < uniqueRowValues.Count; k++) {
							totalNumUniqueRowValues += uniqueRowValues[k].count;
						}
						//Add a point corresponding to the percentage of this specific unique row value relative to the total number of values. Includes unknown rows.
						pieSeries.Points.AddXY(xValues[j], Math.Round((((float)uniqueRowValues[j].count) / ((float)totalNumUniqueRowValues + columnValueCounters[0].unknownCount)) * 100, 1));
					}
				}
			}

			//Reorders the data to be in accordance with the user specified order of the data.
			ApplyItemOrder(pieSeries);

			//Add another column that represents the number of "Unknown" responses (rows that were blank in that column).
			if (options.isCount) {
				pieSeries.Points.AddXY("Unknown", columnValueCounters[0].unknownCount);
			}
			else if (options.isPercentage) {
				pieSeries.Points.AddXY("Unknown", Math.Round((((float)columnValueCounters[0].unknownCount) / ((float)totalNumUniqueRowValues + columnValueCounters[0].unknownCount)) * 100, 1));
			}
			//Apply all of the customized settings for the series and add it to the chart.
			FinalizePieChartSeries(chart, pieSeries, columnValueCounters, options);
			//Apply all of the customized settings for the chart itself.
			ApplyCustomChartOptions(chart, options, columnValueCounters);

			try {
				chart.SaveImage(filename, ChartImageFormat.Png);
			}
			catch (Exception) {
				throw new Exception("Failed to create graph image");
			}
			return filename;
		}
		#region ApplyItemOrder
		///<summary>Reorders the data in the given series to be in accordance with the user specified order of the data.</summary>
		public void ApplyItemOrder(Series pieSeries) {
			//Reorders the data to be in accordance with the user specified order of the data.
			if (_itemOrder != null) {
				for (int i = 0; i < pieSeries.Points.Count; i++) {
					DataPoint point = pieSeries.Points[i];
					if (_itemOrder.Contains(point.AxisLabel)) {
						DataPoint temp = point;
						pieSeries.Points[i] = pieSeries.Points[_itemOrder.IndexOf(point.AxisLabel)];
						pieSeries.Points[_itemOrder.IndexOf(point.AxisLabel)] = temp;
					}
				}
			}
		}
		///<summary>Reorders the data in all of the given series' to be in accordance with the user specified order of the data.</summary>
		public void ApplyItemOrder(List<Series> allSeries) {
			if (_itemOrder != null) {
				for (int i = 0; i < allSeries.Count; i++) {
					string seriesName = allSeries[i].Name;
					//If the user specified one of these 
					if (_itemOrder.Contains(seriesName)) {
						Series temp = allSeries[i];
						allSeries[i] = allSeries[_itemOrder.IndexOf(seriesName)];
						allSeries[_itemOrder.IndexOf(seriesName)] = temp;
					}
				}
			}
		}
		#endregion
		#endregion
		#region Finalize Bar and Pie Chart Series
		/// <summary>
		/// Configures all style settings for an individual bar chart data series.
		/// </summary>
		private void FinalizeBarChartSeries(Chart chart, Series series, List<ColumnValueCounter> columnValueCounters, Color color, DocumentManipulation.TextReplacementOptions options) {
			series.ChartType = SeriesChartType.Column;
			series["PieLabelStyle"] = "Outside";
			series.Color = color;
			series.IsVisibleInLegend = true;
			//Set series label style
			series.CustomProperties = "BarLabelStyle = Top";
			series.CustomProperties = "LabelStyle = Top";
			series.Font = new System.Drawing.Font("Calibri", 16);
			series.IsValueShownAsLabel = true;
			if (options.isPercentage) {
				foreach (DataPoint point in series.Points) {
					if (!point.YValues[0].ToString().Contains('.'))
						point.Label = point.YValues[0].ToString() + ".0";
				}
			}
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
		private void FinalizePieChartSeries(Chart chart, Series series, List<ColumnValueCounter> columnValueCounters, DocumentManipulation.TextReplacementOptions options) {
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
			}
			if (options.isPercentage) {
				foreach (DataPoint point in series.Points) {
					if (!point.YValues[0].ToString().Contains('.'))
						point.Label = point.YValues[0].ToString() + ".0";
				}
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
			chart.ChartAreas["main"].AxisY.MinorGrid.Interval = 2;
			chart.ChartAreas["main"].AxisY.MajorGrid.LineColor = Color.Gray;
			chart.ChartAreas["main"].AxisY.MajorGrid.Interval = 2;
			chart.ChartAreas["main"].AxisY.Interval = 2;
			if (options.isPercentage) {
				chart.ChartAreas["main"].AxisY.Maximum = 100;
				chart.ChartAreas["main"].AxisY.MinorGrid.Interval = 10;
				chart.ChartAreas["main"].AxisY.MajorGrid.Interval = 10;
				chart.ChartAreas["main"].AxisY.Interval = 10;
				foreach (Series series in chart.Series) {
					foreach (DataPoint point in series.Points) {
						string pointLabel = point.YValues[0].ToString();
						if (!pointLabel.Contains('.')) {
							point.YValues[0] = Math.Round((float)point.YValues[0] + 0.00001f, 1);
						}
					}
				}
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
				legend.LegendItemOrder = LegendItemOrder.ReversedSeriesOrder;
				chart.Legends.Add(legend);
			}
			if (options.graphType == "pie") {
				legend.LegendStyle = LegendStyle.Table;
				legend.Docking = Docking.Right;
				legend.Font = new System.Drawing.Font("Calibri", 16);
				legend.IsEquallySpacedItems = true;
				legend.TableStyle = LegendTableStyle.Wide;
				chart.Series[0].IsVisibleInLegend = false;
				int i = 0;
				for (; i < columnValueCounters[0].uniqueRowValues.Count; i++) {
					LegendItem legendItem = new LegendItem();
					legendItem.Name = columnValueCounters[0].uniqueRowValues[i].name;
					legendItem.Color = chart.Series[0].Points[i].Color;
					legendItem.BorderColor = Color.Transparent;
					legendItem.MarkerBorderColor = Color.Transparent;
					legend.CustomItems.Add(legendItem);
				}
				if (columnValueCounters[0].uniqueRowValues.Count < chart.Series[0].Points.Count) {
					LegendItem legendItem = new LegendItem();
					legendItem.Name = "Unknown";
					legendItem.Color = chart.Series[0].Points[i].Color;
					legendItem.BorderColor = Color.Transparent;
					legendItem.MarkerBorderColor = Color.Transparent;
					legend.CustomItems.Add(legendItem);
				}
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
		/// <summary>
		/// Builds and returns a string based on the data in the columns that were specified by the user's command as well as the options the user specified.
		/// </summary>
		public string GenerateText(List<ColumnValueCounter> usedColumns, DocumentManipulation.TextReplacementOptions processedCommand) {
			string assembledText = "";
			int unknownCount;
			//If the user command specified only a column name then return the name of the first value in the first used column.
			if (processedCommand.isColumnValue) {
				return usedColumns[0].uniqueRowValues[0].name;
			}
			//If the user command specified the 'count' option, use the count of the number of occurences in each row 
			//value as well as the names of the values to output a string that contains all of the relevant data counts.
			else if (processedCommand.isCount) {
				//Look at each column
				foreach (ColumnValueCounter column in usedColumns) {
					unknownCount = column.totalColumnValues;
					//and look through each row in each column
					foreach (UniqueRowValue row in column.uniqueRowValues) {
						//if there is more than one type of value in the column then format the output string appropriately.
						if (column.uniqueRowValues.Count > 1) {
							assembledText += row.name;
							assembledText += ": ";
							assembledText += row.count;
							assembledText += ", ";
						}
						//Otherwise just sum up the count
						else {
							assembledText += row.count;
						}
						unknownCount -= row.count;
					}
					//Account for any blank rows in each column.
					if (unknownCount != 0) {
						assembledText += "Unknown: ";
						assembledText += unknownCount;
					}
				}
			}
			//If the user command specified a range, find the highest value and the lowest value to output as a range.
			else if (processedCommand.isRange) {
				int lowest = int.MaxValue;
				int highest = int.MinValue;
				int current;
				//Iterate through each row in each column and check the integer value of each row value to find the max and min of that set of values.
				foreach (ColumnValueCounter column in usedColumns) {
					foreach (UniqueRowValue row in column.uniqueRowValues) {
						if ((current = WordToInt(row.name)) != -1) {
							if (current < lowest)
								lowest = current;
							if (current > highest)
								highest = current;
						}
						else {
							throw new Exception("Attempted to make range from invalid data input.");
						}
					}
				}
				//Create a string representation of the range of values
				assembledText += lowest;
				assembledText += " - ";
				assembledText += highest;
			}
			//If the user command specified the 'mean' option, calculate the mean of all of the values and return it as a string.
			else if (processedCommand.isMean) {
				int total = 0;
				int current = 0;
				int uniqueRowValueCount = 0;
				//sum up the integer values of each row value in each column
				foreach (ColumnValueCounter column in usedColumns) {
					foreach (UniqueRowValue row in column.uniqueRowValues) {
						if ((current = WordToInt(row.name)) != -1) {
							total += (current * row.count);
						}
						uniqueRowValueCount += row.count;
					}
				}
				//divide by the total number of values to get the mean and round to 2 decimal places.
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
