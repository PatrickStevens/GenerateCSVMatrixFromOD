# GenerateCSVMatrixFromOD

Generate a comma separated file that represents a Network Analyst Origin-Destination Cost Matrix analysis result.

### How to use the application

From a windows command line, call the executable and pass in the path to your input Origin-Destination Cost Matrix layer file, like this:

> GenerateCSVMatrixFromOD C:\MyInputODLayer.lyr

The output CSV files will be stored next to the layer file.

### Output

There are two CSVs output:

* OD_ODCostFlatTable_OptimizedOn_&lt;Impedance name&gt;
    * Output each row of the ODLines as Origin,Destination,&lt;Impedance name&gt;
* OD_ODCostMatrix_OptimizedOn_&lt;Impedance name&gt;
    * Output a true matrix.  Each column is a destination.  Each row is an origin.

### Downloads of executables

https://www.arcgis.com/home/item.html?id=00a6a60a32b54393996f0a8ce9ee02c9
