using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using System.Data;
namespace TemplatingProject {
	/// <summary>
	/// This class handles the collection and agregation of all data. This class supports importing a CSV file of the appropriate format, assembling the data into more usable objects (ColumnValueCounters),
	/// and associating those columns with their respective excel column tags.
	/// </summary>
	public class DataCollection {
		#region ClassWideVariables
		private static DataTable _allData;
		private List<string> _columnHeaders = new List<string>();
		#endregion
		#region ImportCSV
		/// <summary>
		/// Prompts the user to select a CSV file to open, opens that file, and places all of the data from it into a data table to be easily manipulated.
		/// Returns true for success and false for failure.
		/// </summary>
		public bool ImportCSV() {
			//Create a new topmost form to put the file dialog on to make sure it shows up in front of all other windows.
			Form topmostForm = new Form { TopMost = true };
			//Create a new file dialog filtered to only show CSV files
			OpenFileDialog selectFile = new OpenFileDialog {
				Filter = "CSV Files (*.csv)|*.csv",
				//Use legacy windows open file dialog because the new one crashes.
				AutoUpgradeEnabled = false
			};
			DialogResult result = DialogResult.None;
			try {
				result = selectFile.ShowDialog(topmostForm);
			}
			catch (Exception) {
				MessageBox.Show(new Form { TopMost = true }, "Error opening file selection window. Try restarting the program");
				return false;
			}
			//If a file dialog was created and a file was selected successfully, proceed getting the data from that file.
			if (result == DialogResult.OK) {
				string CSVFilePathName = selectFile.FileName;
				DataTable dt = new DataTable();
				try {
					//Insert all of the column headers into the data table
					GetColumns(dt, CSVFilePathName);
					int columnCount = _columnHeaders.Count;
					//Insert all of the rows of data into the data table.
					GetRows(dt, CSVFilePathName);
					_allData = dt;
					return true;
				}
				catch (IOException) {
					MessageBox.Show(new Form { TopMost = true }, "Error: Cannot open CSV file while it is in use by another program", "File Access Error");
					return false;
				}
				catch (Exception e) {
					MessageBox.Show(new Form { TopMost = true }, "Error: " + e.Message);
					return false;
				}
			}
			else {
				return false;
			}
		}
		#endregion
		#region AssembleColumnValueCounters
		public List<ColumnValueCounter> AssembleColumnValueCounters() {
			//columnValueCounters stores a ColumnValueCounter for every column in the data table.
			//This allows for storing the names and number of occurences of each unique data row value in relation to the column that it is a part of.
			List<ColumnValueCounter> columnValueCounters = new List<ColumnValueCounter>();
			//columnHeaders is a list of all of the column headers that we get from the importCSV function
			List<string> columnHeaders = _columnHeaders;
			//For every column, make a new ColumnValueCounter. 
			//For each unique row value in the column, count the number of occurences of that value and store both the value and the count in the ColumnValueCounter UniqueRowValue attribute.
			
			for (int i = 0; i < _allData.Columns.Count; i++) {
				ColumnValueCounter currentColumn = new ColumnValueCounter {
					columnName = columnHeaders[i],
					totalColumnValues = _allData.AsDataView().ToTable(false, columnHeaders[i]).Rows.Count
				};
				DataRowCollection uniqueRows = _allData.AsDataView().ToTable(true, columnHeaders[i]).Rows;
				//for each unique row value in this column
				UniqueRowValue unknownRowValue = new UniqueRowValue("Unknown", CountUniqueRows(_allData, currentColumn.columnName, ""));
				currentColumn.unknownCount = unknownRowValue.count;
				for (int j = 0; j < uniqueRows.Count; j++) {
					//Gets the name of the unique row value
					string uniqueRowName = uniqueRows[j].ItemArray[0].ToString();
					//Exclude unique row values that we do not care about (the column header and any blank rows)

					if (uniqueRowName != "" && uniqueRowName != currentColumn.columnName) {
						currentColumn.uniqueRowValues.Add(new UniqueRowValue(uniqueRowName, CountUniqueRows(_allData, currentColumn.columnName, uniqueRowName)));
					}
					
				}
				//Get abbreviated string representation of the current column (i.e. AA or AF)
				currentColumn.abbreviatedRepresentation = GetExcelColumnName(i);
				//Sorts the row values alphabetically within the current column object.
				currentColumn.uniqueRowValues.Sort((x, y) => x.name.CompareTo(y.name));
				//store that entire column
				columnValueCounters.Add(currentColumn);
			}
			return columnValueCounters;
		}
		#endregion
		#region Fill Data Table
		/// <summary>
		/// Gets the column headers and other information about each column and inserts them into a data table
		/// </summary>
		/// <param name="dt">The data table to put the column values into</param>
		/// <param name="CSVFilePathName">The file path for the CSV file that we read the column values from</param>
		private void GetColumns(DataTable dt, string CSVFilePathName) {

			string[] firstHeaderRow;
			string[] secondHeaderRow;
			using (StreamReader columnReader = new StreamReader(CSVFilePathName)) {
				//Get all of the column header names and place them into the data table as Columns.
				//Data is formatted such that the column headers could be in either the first or second row of the CSV file. The following code performs that filtering.
				firstHeaderRow = columnReader.ReadLine().Split(',');
				secondHeaderRow = columnReader.ReadLine().Split(',');
				for (int i = 0; i < secondHeaderRow.Length; i++) {
					if (secondHeaderRow[i] == null || secondHeaderRow[i].Length == 0) {
						continue;
					}
					//If there are quotes in the column name, trims the quotes off of the headers for better looking graphs when the headers are used to make graphs.
					if (secondHeaderRow[i][0] == '\"') {
						//Split on the " character to get the actual string that we want at index 1.
						secondHeaderRow[i] = secondHeaderRow[i].Split('\"')[1] + secondHeaderRow[i + 1].Split('\"')[0];
						//Shift the columns back into their correct positions after splitting created an extra array entry.
						for (int j = (i + 2); j < secondHeaderRow.Length; j++) {
							secondHeaderRow[j - 1] = secondHeaderRow[j];
						}
						secondHeaderRow[secondHeaderRow.Length - 1] = null;
					}
				}
				for (int i = 0; i < Math.Min(secondHeaderRow.Length, firstHeaderRow.Length); i++) {
					//There are some dummy entries in the second header row that should not be useed as column headers.
					//If the current column entry is not one of these then we do use it as a column header.
					if (secondHeaderRow[i] != "Response" && secondHeaderRow[i] != "" && secondHeaderRow[i] != null && secondHeaderRow[i] != "Open-Ended Response") {
						_columnHeaders.Add(secondHeaderRow[i]);
					}
					//Otherwise we use the entry in the firstHeaderRow as the column header for this column.
					else if (firstHeaderRow[i] != "") {
						_columnHeaders.Add(firstHeaderRow[i]);
					}
				}
				//For each column of data add a column to the data table.
				foreach (string header in _columnHeaders) {
					//Check for duplicate column entries
					if (dt.Columns.Contains(header)) {
						MessageBox.Show(new Form { TopMost = true }, "Warning: Multiple identical column names detected in CSV input");
						dt.Columns.Add(header + "copy");
					}
					else
						dt.Columns.Add(header);
				}
			}
		}
		/// <summary>
		/// Reads data from a CSV file and inserts it row-by-row into a data table object.
		/// </summary>
		/// <param name="dt">The data table to insert into</param>
		/// <param name="CSVFilePathName">The file name of the CSV file</param>
		private void GetRows(DataTable dt, string CSVFilePathName) {
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
					for (int i = 0; i < _columnHeaders.Count; i++) {
						if (rows[i].Length > 0) {
							//Check to see if the row had commas in it that would make the intended row value get split up into multiple row values.
							if (rows[i][0] == '\"') {
								//Remove the quote at the beginning of the first row value
								rows[i] = rows[i].Remove(0, 1);
								int j = i + 1;
								//Loop through following row values, combining them with the original one as we go
								for (; !rows[j].Contains("\""); j++) {
									rows[i] += ("," + rows[j]);
									rows[j] = "COMMAFIX";
								}
								//Remove the quote from the last piece of the row value
								rows[i] += ("," + rows[j].Remove(rows[j].Length - 1, 1));
								rows[j] = "COMMAFIX";
							}
						}
						//Make sure that we are not getting any of the dummy data entries.
						if (rows[i] != "Response" && rows[i] != "Open-Ended Response" && rows[i] != "COMMAFIX") {
							dr[i] = rows[i];
						}
					}
					//add our row to the data table
					dt.Rows.Add(dr);
				}
			}
		}
		#endregion
		#region Utility Functions
		/// <summary>
		/// Gets the string representation of an excel column identifier from the column number.
		/// Column number is an ineger where 0 indexes the first column.
		/// </summary>
		private string GetExcelColumnName(int columnNumber) {
			int dividend = columnNumber + 1;
			string columnName = String.Empty;
			int modulo;

			while (dividend > 0) {
				modulo = (dividend - 1) % 26;
				columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
				dividend = (int)((dividend - modulo) / 26);
			}

			return columnName;
		}
		/// <summary>
		/// Returns the number of unique row values that have a specified name in a column that is specified by name.
		/// </summary>
		private int CountUniqueRows(DataTable allData, string columnName, string uniqueRowName) {
			string filterExpression = "[" + columnName + "]" + " = '" + uniqueRowName + "\'";
			return allData.Select(filterExpression).Length;
		}
		#endregion
	}
	#region Data Containers
	/// <summary>
	/// An object that stores a the name and number of occurences of a particular value in a column.
	/// </summary>
	public struct UniqueRowValue {
		public string name;
		public int count;
		public UniqueRowValue(string name, int count) {
			this.name = name;
			this.count = count;
		}
	}
	/// <summary>
	/// An object that holds information about a column of the data.
	/// </summary>
	public class ColumnValueCounter {
		public List<UniqueRowValue> uniqueRowValues;
		public string columnName;
		public int totalColumnValues;
		public int unknownCount;
		public string abbreviatedRepresentation;
		public ColumnValueCounter() => uniqueRowValues = new List<UniqueRowValue>();
	}
	#endregion
}
