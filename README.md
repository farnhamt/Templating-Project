Documentation for Word Document Templater Application:

PURPOSE:
The purpose of this application is to allow the generation of word document reports quickly based on CSV file input and a
word document template file. It was created for the 4-H program at Oregon State University to save time when executing the
task of creating reports from data that their program collects via Qualtrics surveys. The project allows for the creation of
future templates by the user using a simple word document command syntax that is defined below.


TO USE THIS PROGRAM:
0. Making a backup of all files involved is recommended. If the program is interrupted abruptly it may overwrite file data.
1. Create a word document template using the instructions below on writing template commands.
2. Create a CSV file as input that follows the appropriate format (qualtrics output format).
3. Run 'Templater.exe'.
4. First, select the word document that you want to use as a template.
5. Second, select the .CSV file that you want to use as data input.
6. Third, select or name the .docx file that you would like to save the resulting document (do not overwrite your template!)


WRITING TEMPLATE COMMANDS:

Commands are specified in the document by surrounding them in triple curly brace characters. (i.e. {{{command}}}})

There are many different options that are allowed in each command. These options are seperated by semi-colon characters.
The segments ignore leading and trailing space characters.

The following options are available in the commands:

-The first option in the command is always the command execution type. This option can have a value of "text", "bar", or "pie".
The command execution type designates whether the command is generating text, generating a bar chart, or generating a pie chart.

-The second command option specifies how the data will be interpreted when generating the output of the command (bar chart, pie chart, or text);
The second command can have the following values: "range", "mean", "percentage" or "%", or "count".
The "range" and "mean" options are only available when the command execution type is set to "text". "percentage" and "count" options 
are the only ones that can be used on graphs.
If no data interpretation option is given, the command will default to providing the first value in the data column specified in following options.
This is the last option that is specified for the "text" execution type.

-The following options have multiple different cases for what they can be. The third option can be either an x-axis font size or a CSV Column name specification.
If the value of this option is a number then the templater will interpret this option as a manually provided font size for the x-axis labels on the graph.
Otherwise, this option marks the beginning of the column name specifications.

-There can be any number of column name specification options in a row. These column names represent the columns that will be used to generate the graph.
The column names can be provided as either their excel column heading (A, L, AD, BF, etc.) or the actual title of the column ("Are you thinking about getting a job in the year after school?").
If using the actual title of the column, the column name needs to be specified exactly as it appears in the CSV file.

-The last option that is provided is the title of the graph (i.e. "Percentage of Youth Reporting"). This is always the last option provided.

-There are two special command cases that are used to provide imporant user control.
The first of these commands declares the color pallette that is used in the generated graphs.
The color pallette is declared with either the keyword "colors" or the keyword "colorpallette" and the command uses the following format: 
{{{colors; r1,g1,b1; r2,g2,b2; r3,g3,b3; ... ; rn,gn,bn}}}
where the r,g, and b values represent the rgb values in a 0-255 RGB color specification. The colors will be used in the graphs in the order that they appear in the specified color pallette.
An example of a color pallette specification command is as follows:
{{{colors; 215, 63, 9; 183, 169, 154; 170, 157, 46; 74, 119, 60}}}

The second of these special commands dictates the order in which values should be displayed in a graph.
This command can be accessed using the keyword "order" in a command that may look like {{{order;Yes;No}}}.
The above command would place the column that measures "Yes" data values in the first graph index and the column that measures "No" data values in the second graph index.
The values passed into the 'order' command must be identical to their corresponding values in the CSV input file and are case sensitive.
The order command also changes the order in which the values show up in the graph legend.

Both the 'order' command and the 'colors' command set persistent values that will be used for all further execution of the program until otherwise overwritten by another 'order' or 'colors' command respectively.

The templater scans the document from top to bottom, so different color pallettes and orders can be specified in various spots 
throughout the word document template to ensure that the color scheme for individual graphs is accurate.
Specify the colors in the color pallette in a specific order so they match the intended colors of the values that are represented in the graphs.

-A full command for generating a bar graph might look like this:
{{{bar;%; R; S; T; Participate in a mock interview?; Talk about how to have a professional image on social media?; Percentage of Youth Reporting}}}
The above command specifies a bar graph that analyzes each value within specified columns as a percentage value. It uses CSV columns R, S, T, and
the columns with titles "Participate in a mock interview?" and "Talk about how to have a professional image on social media?".
The title of this graph is given as "Percentage of Youth Reporting".


CSV FILE FORMAT:

The format for the CSV input files is that of standard qualtrics survey output.
The first two rows in the file are used as column headers. The top row is used for primary headers, and the second row is used for secondary headers.
The secondary header row is checked first for each column to see if it is valid by ruling out certain values ("Open-Ended Response", "Response", "", etc.)
If the secondary header row is valid for a data column then the primary header row for that column is not checked and the secondary header is used for that column.
If the secondary header row is not valid then the primary header row for that column will be checked for validity and will throw an error if it is not valid.
All rows below the primary and secondary header rows are treated as data rows.

In the data rows, all values are accepted. Blank data row values are treated as "Unknown" values and will be outputted as such in graph and text creation.
If a column contains numerical values in plain english form ("one", "two", "Three", etc.) they will be treated as their corresponding integers instead (only handles ints 0-20).