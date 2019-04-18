using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Charting = System.Windows.Forms.DataVisualization.Charting;
using Chart = System.Windows.Forms.DataVisualization.Charting.Chart;
using Series = System.Windows.Forms.DataVisualization.Charting.Series;
using Color = System.Drawing.Color;

namespace TemplatingProject {
	class DocumentCommandExecuter {
		private Color[] _colorPallette;
		public DocumentCommandExecuter(Color[] colorPallette) {
			_colorPallette = colorPallette;
		}
		#region GenerateGraph
		public string GenerateGraph(List<ColumnValueCounter> columnValueCounters, string filename, DocumentManipulation.TextReplacementOptions options, Word.Application wordApp) {
			//Need to get the data that needs to be in the graph here.
			//allData = Form1.allData
			if (options.graphType == "bar") {
				return this.GenerateBarGraph(columnValueCounters, filename, options);
			}
			else if (options.graphType == "pie") {
				return this.GeneratePieChart(columnValueCounters, filename, options);
			}
			return filename;
		}
		#endregion
		#region GenerateBarGraph
		private string GenerateBarGraph(List<ColumnValueCounter> columnValueCounters, string filename, DocumentManipulation.TextReplacementOptions options) {

			Chart chart = new Chart();
			int numSeries = options.columnNames.Count();

			List<Charting.Series> allSeries = new List<Charting.Series>();
			string[] xValues;
			if (columnValueCounters.Count == 0) {
				MessageBox.Show("Error: No data imported. Invalid CSV data format");
				Environment.Exit(1);
			}
			if (columnValueCounters.Count > 1) {
				xValues = new string[columnValueCounters.Count];
				for (int i = 0; i < columnValueCounters[0].uniqueRowValues.Count; i++) {
					allSeries.Add(new Charting.Series(columnValueCounters[0].uniqueRowValues[i].name));
				}
				for (int i = 0; i < columnValueCounters.Count; i++) {
					xValues[i] = columnValueCounters[i].columnName;
				}
				CreateFillerChartSeries(chart, "beginning", columnValueCounters, xValues.ToList());
				CreateFillerChartSeries(chart, "beginning1", columnValueCounters, xValues.ToList());
			}
			else {
				if (columnValueCounters[0].unknownCount > 0) {
					xValues = new string[columnValueCounters[0].uniqueRowValues.Count + 1];
					xValues[columnValueCounters[0].uniqueRowValues.Count] = "Unknown";
				}
				else {
					xValues = new string[columnValueCounters[0].uniqueRowValues.Count];
				}
				allSeries.Add(new Charting.Series("0"));
				for (int i = 0; i < columnValueCounters[0].uniqueRowValues.Count; i++) {
					xValues[i] = columnValueCounters[0].uniqueRowValues[i].name;
				}
			}

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
						CreateFillerChartSeries(chart, allSeries[j].Name, columnValueCounters, xValues.ToList());
					}
					//Apply all of the customized settings for the series and then add the series to the chart.
					FinalizeBarChartSeries(chart, allSeries[j], columnValueCounters, _colorPallette[j]);
				}
				CreateFillerChartSeries(chart, "end", columnValueCounters, xValues.ToList());
				CreateFillerChartSeries(chart, "end1", columnValueCounters, xValues.ToList());
			}
			//Apply all of the customized settings for the chart itself.
			ApplyCustomChartOptions(chart, options, columnValueCounters);

			try {
				chart.SaveImage(filename, Charting.ChartImageFormat.Png);
			}
			catch (Exception e) {
				MessageBox.Show("Error: Failed to create graph image");
			}
			return filename;
		}
		#endregion
		#region GeneratePieChart
		private string GeneratePieChart(List<ColumnValueCounter> columnValueCounters, string filename, DocumentManipulation.TextReplacementOptions options) {

			Chart chart = new Chart();
			int numSeries = options.columnNames.Count();

			List<Charting.Series> allSeries = new List<Charting.Series>();
			string[] xValues;
			if (columnValueCounters.Count > 1) {
				xValues = new string[columnValueCounters.Count];
				for (int i = 0; i < columnValueCounters[0].uniqueRowValues.Count; i++) {
					allSeries.Add(new Charting.Series(columnValueCounters[0].uniqueRowValues[i].name));
				}
				for (int i = 0; i < columnValueCounters.Count; i++) {
					xValues[i] = columnValueCounters[i].columnName;
				}
			}
			else {
				if (columnValueCounters[0].unknownCount > 0) {
					xValues = new string[columnValueCounters[0].uniqueRowValues.Count + 1];
					xValues[columnValueCounters[0].uniqueRowValues.Count] = "Unknown";
				}
				else {
					xValues = new string[columnValueCounters[0].uniqueRowValues.Count];
				}
				allSeries.Add(new Charting.Series("0"));
				for (int i = 0; i < columnValueCounters[0].uniqueRowValues.Count; i++) {
					xValues[i] = columnValueCounters[0].uniqueRowValues[i].name;
				}
			}

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
						CreateFillerChartSeries(chart, allSeries[j].Name, columnValueCounters, xValues.ToList());
					}
					//Apply all of the customized settings for the series and then add the series to the chart.
					FinalizeBarChartSeries(chart, allSeries[j], columnValueCounters, _colorPallette[j]);
				}
			}
			//Apply all of the customized settings for the chart itself.
			this.ApplyCustomChartOptions(chart, options, columnValueCounters);

			try {
				chart.SaveImage(filename, Charting.ChartImageFormat.Png);
			}
			catch (Exception e) {
				MessageBox.Show("Error: Failed to create graph image");
			}
			return filename;
		}
		#endregion
		#region Finalize Bar and Pie Chart Series
		private void FinalizeBarChartSeries(Chart chart, Series series, List<ColumnValueCounter> columnValueCounters, Color color) {
			series.ChartType = Charting.SeriesChartType.Column;
			series["PieLabelStyle"] = "Outside";
			series.IsValueShownAsLabel = true;
			series.SmartLabelStyle.Enabled = false;
			series.Color = color;
			series.IsVisibleInLegend = true;

			series.CustomProperties = "BarLabelStyle = Top";
			series.CustomProperties = "LabelStyle = Top";
			series.Font = new System.Drawing.Font("Calibri", 16);
			series.SmartLabelStyle.AllowOutsidePlotArea = Charting.LabelOutsidePlotAreaStyle.Partial;
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

		private void FinalizePieChartSeries(Chart chart, Series series, List<ColumnValueCounter> columnValueCounters) {
			series.ChartType = Charting.SeriesChartType.Pie;
			//Sort points by largest y value (percentage or count) to avoid the pie chart slices appearing in random order
			for (int i = 0; i < series.Points.Count; i++) {
				Charting.DataPoint temp = new Charting.DataPoint();
				for (int j = i; j < series.Points.Count; j++) {
					if (series.Points[j].YValues[0] > series.Points[i].YValues[0]) {
						temp = series.Points[i];
						series.Points[i] = series.Points[j];
						series.Points[j] = temp;
					}
				}
				series.Points[i].IsVisibleInLegend = true;
			}
			if (series.Points.Count > _colorPallette.Length) {
				MessageBox.Show("Not enough colors in color palette to satisfy all data points");
				for (int i = 0; i < series.Points.Count; i++) {
					series.Points[i].Color = _colorPallette[i % _colorPallette.Length];
				}
			}
			else {
				for (int i = 0; i < series.Points.Count; i++) {
					series.Points[i].Color = _colorPallette[i];
				}
			}
			series.IsValueShownAsLabel = true;
			series.SmartLabelStyle.Enabled = true;
			//series.Color = color;
			series.IsVisibleInLegend = true;
			series.CustomProperties = "BarLabelStyle = Top";
			series.CustomProperties = "LabelStyle = Top";
			series.Font = new System.Drawing.Font("Calibri", 16);
			chart.Series.Add(series);
		}
		#endregion
		#region CreateFillerChartSeries
		private void CreateFillerChartSeries(Chart chart, string fillerName, List<ColumnValueCounter> columnValueCounters, List<string> xValues) {
			if (columnValueCounters.Count > 1) {
				Series filler = new Series("filler" + fillerName);
				foreach (string xValue in xValues) {
					filler.Points.AddXY(xValue, 0);
				}
				filler.Enabled = true;
				filler.IsVisibleInLegend = false;
				filler.SetCustomProperty("PointWidth", "0.1");
				chart.Series.Add(filler);
			}
		}
		#endregion
		#region applyCustomChartOptions
		private void ApplyCustomChartOptions(Chart chart, DocumentManipulation.TextReplacementOptions options, List<ColumnValueCounter> columnValueCounters) {
			Charting.ChartArea chartArea = new Charting.ChartArea("main");
			Charting.Axis yAxis = new Charting.Axis(chartArea, Charting.AxisName.Y);
			Charting.Axis xAxis = new Charting.Axis(chartArea, Charting.AxisName.X);
			chart.ChartAreas.Add(chartArea);
			//Set Y-Axis Style
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
			chart.ChartAreas["main"].AxisY.MajorGrid.LineColor = System.Drawing.Color.LightGray;
			chart.ChartAreas["main"].AxisY.LineWidth = 0;
			chart.ChartAreas["main"].AxisY.LabelAutoFitMinFontSize = 16;
			chart.ChartAreas["main"].AxisY.LabelAutoFitMaxFontSize = 16;
			//Set X-Axis style
			chart.ChartAreas["main"].AxisX.MajorTickMark.Enabled = false;
			chart.ChartAreas["main"].AxisX.MajorGrid.Enabled = false;
			chart.ChartAreas["main"].AxisX.LineWidth = 0;
			//Set x-axis label style
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

			chart.BorderlineDashStyle = Charting.ChartDashStyle.Solid;
			chart.BorderlineColor = System.Drawing.Color.LightGray;
			chart.BorderlineWidth = 1;
			chart.AntiAliasing = Charting.AntiAliasingStyles.None;
			//Configure graph legend
			Charting.Legend legend = new Charting.Legend();
			legend.Font = new System.Drawing.Font("Calibri", 14);
			legend.IsTextAutoFit = false;
			
			legend.Alignment = System.Drawing.StringAlignment.Center;
			legend.LegendStyle = Charting.LegendStyle.Row;
			if (columnValueCounters.Count > 1) {
				legend.Docking = Charting.Docking.Bottom;
				chart.ChartAreas["main"].AxisX.LabelAutoFitMaxFontSize = 16;
				chart.Legends.Add(legend);
			}
			if (options.graphType == "pie") {
				legend.LegendStyle = Charting.LegendStyle.Table;
				legend.IsEquallySpacedItems = true;
				legend.Font = new System.Drawing.Font("Calibri", 16);
				chart.Series[0].Font = new System.Drawing.Font("Calibri", 16);
				chart.Legends.Add(legend);
			}

			Charting.Title title = new Charting.Title();
			title.Text = options.graphTitle;
			title.Font = new System.Drawing.Font("Calibri", 24, System.Drawing.FontStyle.Italic);
			title.ForeColor = Color.Gray;
			chart.Titles.Add(title);
			if (options.graphType == "pie") {
				chart.BorderlineWidth = 0;
				chart.Width = 840;
				chart.Height = 700;
			}
			else if (columnValueCounters.Count > 1) {
				chart.Width = 1200;
				chart.Height = 600;
			}
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
						if ((current = this.WordToInt(row.name)) != -1) {
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
						if ((current = this.WordToInt(row.name)) != -1) {
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
		#region WordToInt
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
