using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.NetworkAnalyst;

namespace GenerateCSVMatrixFromOD
{
  public static class GenerateCsvMatrixFromOD
  {
    public static void Main(string[] args)
    {
      try
      {
        if (args == null || args.Length != 1) throw new ArgumentException("Command line format: <layer file name>");

        // The output is going to be put at the same path as the input layer file.
        var layerFilePath = args[0];
        if (!File.Exists(layerFilePath)) throw new FileNotFoundException($"Layer file does not exist: {layerFilePath}");

        // Set up the licensing.
        if (!ESRI.ArcGIS.RuntimeManager.Bind(ESRI.ArcGIS.ProductCode.EngineOrDesktop))
          throw new ArgumentException("Bind failure.  Is ArcGIS Desktop or Engine installed?");

        IAoInitialize aoInit = new AoInitialize();
        InitializeLicensing(aoInit);

        var context = GetContextFromLayer(layerFilePath);

        List<string> originNames;
        List<string> destinationNames;
        IStringArray costAttributeNames;

        var costMatrixByCostAttributesList = GetCostMatrixByCostAttributesList(context, out costAttributeNames, out originNames, out destinationNames);

        OutputODCostsToCsvFiles(layerFilePath, costMatrixByCostAttributesList, costAttributeNames, originNames, destinationNames);

        aoInit.Shutdown();
      }
      catch (Exception e)
      {
        // Write errors out to the error stream.
        Console.Error.WriteLine("Exception thrown: " + e.Message);
        throw;
      }
    }

    private static void InitializeLicensing(IAoInitialize aoInit)
    {
      // Check out the lowest available license level.
      var licenseLevels = new List<esriLicenseProductCode>()
                {
                    esriLicenseProductCode.esriLicenseProductCodeBasic,
                    esriLicenseProductCode.esriLicenseProductCodeEngine,
                    esriLicenseProductCode.esriLicenseProductCodeStandard,
                    esriLicenseProductCode.esriLicenseProductCodeAdvanced,
                    esriLicenseProductCode.esriLicenseProductCodeArcServer
                };

      var licenseStatus = esriLicenseStatus.esriLicenseUnavailable;
      foreach (var licenseLevel in licenseLevels)
      {
        licenseStatus = aoInit.Initialize(licenseLevel);
        if (licenseStatus == esriLicenseStatus.esriLicenseCheckedOut || licenseStatus == esriLicenseStatus.esriLicenseAlreadyInitialized)
          break;
      }

      if (licenseStatus != esriLicenseStatus.esriLicenseAlreadyInitialized && licenseStatus != esriLicenseStatus.esriLicenseCheckedOut)
        throw new ArcGISLicensingException($"Product license initialization failed. License status: {licenseStatus}");

      licenseStatus = aoInit.CheckOutExtension(esriLicenseExtensionCode.esriLicenseExtensionCodeNetwork);
      if (licenseStatus != esriLicenseStatus.esriLicenseAlreadyInitialized && licenseStatus != esriLicenseStatus.esriLicenseCheckedOut)
        throw new ArcGISLicensingException($"Network extension license initialization failed. License status: {licenseStatus}");
    }

    private static INAContext GetContextFromLayer(string layerFileName)
    {
      ILayerFile layerFile = new LayerFileClass();
      try
      {
        layerFile.Open(layerFileName);
      }
      catch (Exception e)
      {
        throw new InvalidOdLayerException($"Unable to open layer file: {layerFileName}. Make sure it exists and is valid.  Try opening it in ArcMap.{Environment.NewLine}" +
                                          $"Error type: {e.GetType()}{Environment.NewLine}" +
                                          $"Error message: {e.Message}");
      }

      var naLayer = layerFile.Layer as INALayer;
      if (naLayer == null) throw new InvalidOdLayerException($"Unable to get INALayer from layer file: {layerFileName}. Is your layer a Network Analysis layer?");

      var context = naLayer.Context;
      if (context == null) throw new InvalidOdLayerException($"Null context on layer file: {layerFileName}. Is your layer a valid Network Analysis layer?");

      var odSolver = context.Solver as INAODCostMatrixSolver2;
      if (odSolver == null) throw new InvalidOdLayerException($"Input layer '{layerFileName}' must be an Origin-Destination Cost Matrix layer.");

      return context;
    }

    private static List<double[,]> GetCostMatrixByCostAttributesList(
        INAContext context,
        out IStringArray costAttributeNames,
        out List<string> originNames,
        out List<string> destinationNames)
    {
      var odSolver = context.Solver as INAODCostMatrixSolver2;
      if (odSolver == null) throw new InvalidOdLayerException("Unable to get INAODCostMatrixSolver2 from layer file context solver. Is your layer an OD layer?");

      var odCostMatrix = context.Result as INAODCostMatrix;

      // If an matrix already exists in this layer, use it.  If not, solve and generate a matrix.
      if (odCostMatrix == null || !context.Result.HasValidResult || odSolver.MatrixResultType == esriNAODCostMatrixType.esriNAODCostMatrixNone || odSolver.PopulateODLines)
      {
        odCostMatrix = Solve(context, odSolver);
      }

      //////////////////////////////////////////////////////////////////
      //
      //  Set up mapping of OIDs to internal matrix indexes
      //
      //  The reason for doing this is that the internal matrix is stored in such a way that
      //  origins/destinations that map to the same location on the network are stored at the
      //  same index in the matrix. This is an optimization to avoid storing redundant entries
      //  in the matrix.  
      //

      var originMapping = GetInternalIndexToOidMapping(context, odCostMatrix, isOrigin: true);
      var destinationMapping = GetInternalIndexToOidMapping(context, odCostMatrix, isOrigin: false);

      // There are values stored for the impedance, as well as every accumulated attribute.
      costAttributeNames = odCostMatrix.CostAttributeNames;
      var attributeCount = costAttributeNames.Count;
      if (attributeCount == 0)
        throw new InvalidOdLayerException("At least one attribute is required.");

      originNames = new List<string>();
      foreach (var internalOrigins in originMapping)
        originNames.AddRange(internalOrigins.Value);

      destinationNames = new List<string>();
      foreach (var internalDestinations in destinationMapping)
        destinationNames.AddRange(internalDestinations.Value);

      var originCount = originNames.Count;
      var destinationCount = destinationNames.Count;
      if ((originCount == 0) || (destinationCount == 0))
        throw new InvalidOdLayerException("At least one origin and destination are required.");

      //////////////////////////////////////////////////////////////////
      //
      // Iterate the internal matrix to create matrix of cost values by attribute origin and destination
      //

      var matrixUniqueOriginLocationCount = odCostMatrix.OriginCount;
      var matrixUniqueDestinationLocationCount = odCostMatrix.DestinationCount;

      // The index could start at a value less than zero, to represent non-located locations.
      //  Therefore, find the lowest index value before doing the iterations.
      var startIndexMatrixUniqueOriginLocation = GetStartingIndex(originMapping);
      var startIndexMatrixUniqueDestinationLocation = GetStartingIndex(destinationMapping);

      var costMatrixByCostAttributesList = new List<double[,]>();

      for (var attributeIndex = 0; attributeIndex < attributeCount; ++attributeIndex)
      {
        var costs = new double[originCount, destinationCount];
        costMatrixByCostAttributesList.Add(costs);

        var defaultValue = odCostMatrix.DefaultValue[attributeIndex];

        // The outer loop iterates over all of the origins in the internal matrix.
        //  Start at -1 to handle the non-located origins.
        for (int indexMatrixUniqueOriginLocation = startIndexMatrixUniqueOriginLocation, originIndex = 0;
            indexMatrixUniqueOriginLocation < matrixUniqueOriginLocationCount;
            ++indexMatrixUniqueOriginLocation)
        {
          var duplicateOriginCount = originMapping[indexMatrixUniqueOriginLocation].Count;
          for (var duplicateOriginIndex = 0; duplicateOriginIndex < duplicateOriginCount; ++duplicateOriginIndex, ++originIndex)
          {
            for (int indexMatrixUniqueDestinationLocation = startIndexMatrixUniqueDestinationLocation, destinationIndex = 0;
                indexMatrixUniqueDestinationLocation < matrixUniqueDestinationLocationCount;
                ++indexMatrixUniqueDestinationLocation)
            {
              var duplicateDestinationCount = destinationMapping[indexMatrixUniqueDestinationLocation].Count;
              for (var duplicateDestinationIndex = 0; duplicateDestinationIndex < duplicateDestinationCount; ++duplicateDestinationIndex, ++destinationIndex)
              {
                // If either the origin or destination is non-located, use the default value.
                var value = defaultValue;
                if ((indexMatrixUniqueOriginLocation >= 0) && (indexMatrixUniqueDestinationLocation >= 0))
                  value = odCostMatrix.Value[indexMatrixUniqueOriginLocation, indexMatrixUniqueDestinationLocation, attributeIndex];

                costs[originIndex, destinationIndex] = value;
              }
            }
          }
        }
      }

      return costMatrixByCostAttributesList;
    }

    private static INAODCostMatrix Solve(INAContext context, INAODCostMatrixSolver2 odSolver)
    {
      // Set the solver to generate a matrix.
      odSolver.MatrixResultType = esriNAODCostMatrixType.esriNAODCostMatrixFull;
      odSolver.PopulateODLines = false;

      // Solve
      IGPMessages gpMessages = new GPMessagesClass();
      try
      {
        context.Solver.Solve(context, gpMessages, null);
      }
      catch (Exception e)
      {
        var gpMessagesString = GenerateStringOfAllGpMessages(gpMessages);
        throw new InvalidOdLayerException($"Solve exception message: {e.Message}{Environment.NewLine}" +
                                          $"Returned GPMessages are as follows:{Environment.NewLine}{gpMessagesString}");
      }

      var odCostMatrix = context.Result as INAODCostMatrix;
      if (!context.Result.HasValidResult || odCostMatrix == null)
        throw new InvalidOdLayerException("Valid result could not be generated. Is something wrong with your layer? Try solving it in ArcMap.");

      return odCostMatrix;
    }

    private static void OutputODCostsToCsvFiles(
        string layerFilePath,
        IReadOnlyList<double[,]> costMatrixByCostAttributesList,
        IStringArray costAttributeNames,
        List<string> originNames,
        List<string> destinationNames)
    {
      var baseOutputFileName = Path.GetFileNameWithoutExtension(layerFilePath);
      var outputFolderPath = Directory.GetParent(layerFilePath).FullName;

      // For the impedance attribute (1st cost attribute), output costs as a flat table CSV text file.
      // For all cost attributes (impedance and accumulated), output costs as matrix CSV text file.
      //
      var attributeCount = costMatrixByCostAttributesList.Count;
      for (var a = 0; a < attributeCount; ++a)
      {
        var costs = costMatrixByCostAttributesList[a];

        var attributeName = costAttributeNames.Element[a];
        var isImpedance = (a == 0);
        var attributeDescription = $"{(isImpedance ? "OptimizedOn" : "AccumulationOf")}_{attributeName}";

        const string outputFileNameFormat = "{0}_ODCost{1}_{2}.csv";
        string outputFileName;
        string outputFilePath;

        if (isImpedance)
        {
          outputFileName = string.Format(outputFileNameFormat, baseOutputFileName, "FlatTable", attributeDescription);
          outputFilePath = Path.Combine(outputFolderPath, outputFileName);
          OutputCsvFlatTable(outputFilePath, costAttributeNames, originNames, destinationNames, costMatrixByCostAttributesList);
        }

        outputFileName = string.Format(outputFileNameFormat, baseOutputFileName, "Matrix", attributeDescription);
        outputFilePath = Path.Combine(outputFolderPath, outputFileName);


        OutputCsvMatrix(outputFilePath, originNames, destinationNames, costs);
      }
    }

    private static void OutputCsvFlatTable(string outputFilePath, IStringArray costAttributeNames,
      IReadOnlyList<string> originNames, IReadOnlyList<string> destinationNames, IReadOnlyList<double[,]> costMatrixByCostAttributesList)
    {
      if (costAttributeNames == null) throw new ArgumentNullException(nameof(costAttributeNames), "Missing cost attribute names");
      if (originNames == null) throw new ArgumentNullException(nameof(originNames), "Missing origin names");
      if (destinationNames == null) throw new ArgumentNullException(nameof(destinationNames), "Missing destination names");
      if (costMatrixByCostAttributesList == null) throw new ArgumentNullException(nameof(costMatrixByCostAttributesList));

      var sbHeader = new StringBuilder();
      const char cSeparator = ',';
      sbHeader.Append("Origin,Destination");

      var costAttributeCount = costAttributeNames.Count;
      if (costAttributeCount == 0) throw new InvalidOdLayerException("There are no cost attributes");

      var originCount = originNames.Count;
      if (originCount == 0) throw new InvalidOdLayerException("There are no origins");

      var destinationCount = destinationNames.Count;
      if (destinationCount == 0) throw new InvalidOdLayerException("There are no destinations");

      for (var costAttributeIndex = 0; costAttributeIndex < costAttributeCount; ++costAttributeIndex)
      {
        sbHeader.Append(cSeparator);
        var attributeName = NormalizeCsvName(costAttributeNames.Element[costAttributeIndex]);
        sbHeader.Append(attributeName);
      }

      var header = sbHeader.ToString();
      sbHeader.Clear();

      using (var writer = new StreamWriter(outputFilePath, false, Encoding.Unicode))
      {
        writer.WriteLine(header);

        var sbCostsLine = new StringBuilder();

        for (var originIndex = 0; originIndex < originCount; ++originIndex)
        {
          var originName = originNames[originIndex];
          for (var destinationIndex = 0; destinationIndex < destinationCount; ++destinationIndex)
          {
            sbCostsLine.Append(originName);

            var destinationName = destinationNames[destinationIndex];
            sbCostsLine.Append(cSeparator);
            sbCostsLine.Append(destinationName);
            for (var costAttributeIndex = 0; costAttributeIndex < costAttributeCount; ++costAttributeIndex)
            {
              var cost = costMatrixByCostAttributesList[costAttributeIndex][originIndex, destinationIndex];
              sbCostsLine.Append(cSeparator);
              sbCostsLine.Append(cost);
            }

            writer.WriteLine(sbCostsLine.ToString());
            sbCostsLine.Clear();
          }
        }
      }
    }

    private static void OutputCsvMatrix(string outputFilePath, List<string> originNames, List<string> destinationNames, double[,] costs)
    {
      if (string.IsNullOrWhiteSpace(outputFilePath)) throw new ArgumentNullException(nameof(outputFilePath), "Missing output file path");
      if (originNames == null) throw new ArgumentNullException(nameof(originNames), "Missing origin names");
      if (destinationNames == null) throw new ArgumentNullException(nameof(destinationNames), "Missing destination names");
      if (costs == null) throw new ArgumentNullException(nameof(costs), "Missing OD costs");

      var originCount = originNames.Count;
      if (originCount == 0) throw new InvalidOdLayerException("There are no origins");

      var destinationCount = destinationNames.Count;
      if (destinationCount == 0) throw new InvalidOdLayerException("There are no destinations");

      if (costs.GetLength(0) != originCount) return;
      if (costs.GetLength(1) != destinationCount) return;

      var sbHeader = new StringBuilder();
      const char cSeparator = ',';
      sbHeader.Append("Name");

      for (var d = 0; d < destinationCount; ++d)
      {
        sbHeader.Append(cSeparator);
        var destinationName = destinationNames[d];
        sbHeader.Append(destinationName);
      }

      var header = sbHeader.ToString();
      sbHeader.Clear();

      using (var writer = new StreamWriter(outputFilePath, false, Encoding.Unicode))
      {
        writer.WriteLine(header);

        var sbCostsLine = new StringBuilder();

        for (var o = 0; o < originCount; ++o)
        {
          var originName = originNames[o];
          sbCostsLine.Append(originName);

          for (var d = 0; d < destinationCount; ++d)
          {
            var cost = costs[o, d];
            sbCostsLine.Append(cSeparator);
            sbCostsLine.Append(cost);
          }

          writer.WriteLine(sbCostsLine.ToString());
          sbCostsLine.Clear();
        }
      }
    }

    private static int GetStartingIndex<T>(SortedDictionary<int, T> originMapping)
    {
      IEnumerator enumerator = originMapping.Keys.GetEnumerator();
      enumerator.MoveNext();
      return (int)enumerator.Current;
    }

    private static string GenerateStringOfAllGpMessages(IGPMessages gpMessages)
    {
      var gpMessagesStringBuilder = new StringBuilder();
      for (var messageIndex = 0; messageIndex < gpMessages.Count; messageIndex++)
      {
        var gpMessage = gpMessages.GetMessage(messageIndex);
        gpMessagesStringBuilder.AppendLine("Error Code: " + gpMessage.ErrorCode +
                                " Type: " + gpMessage.Type +
                                " Message " + messageIndex + ": " + gpMessage.Description);
      }
      return gpMessagesStringBuilder.ToString();
    }

    private static string NormalizeCsvName(string name)
    {
      return string.IsNullOrEmpty(name) ? null : Regex.Replace(name, @"\s*,\s*", " ");
    }

    /// <summary>
    /// A method to find all of the locations associated with a given internal matrix index.
    /// </summary>
    /// <param name="context"></param>
    /// <param name="odCostMatrix"></param>
    /// <param name="isOrigin"></param>
    /// <returns>
    /// A SortedDictionary of int and List of strings.
    /// The first int is internal matrix ID, the List is names for column and row headers of each location associated with a giving ID.
    /// </returns>
    private static SortedDictionary<int, List<string>> GetInternalIndexToOidMapping(INAContext context, INAODCostMatrix odCostMatrix, bool isOrigin)
    {
      var mapping = new SortedDictionary<int, List<string>>();

      // Get the appropriate NAClass.
      var className = "Origins";
      if (!isOrigin) className = "Destinations";
      var table = context.NAClasses.ItemByName[className] as ITable;
      if (table == null) throw new InvalidOdLayerException($"Unable to get {className} table from NAClasses on the layer context");

      // All of the needed field values.
      var curbApproachFieldIndex = table.FindField("CurbApproach");
      var nameFieldIndex = table.FindField("Name");

      if (curbApproachFieldIndex < 0 || nameFieldIndex < 0)
        throw new InvalidOdLayerException("CurbApproach and Name fields must be present.");

      // Iterate over the specified location table.
      var cursor = table.Search(null, true);
      IRow row;
      while ((row = cursor.NextRow()) != null)
      {
        var naLocationObject = row as INALocationObject;
        var naLocation = naLocationObject?.NALocation;
        if (naLocation == null) continue;

        // Use the NALocation and CurbApproach to find the appropriate internal matrix index.
        var curbApproach = (esriNACurbApproachType)row.Value[curbApproachFieldIndex];
        var internalMatrixIndex = -1;
        // No need to call FindIndex for non-located locations.
        if (naLocation.IsLocated)
        {
          internalMatrixIndex = isOrigin ? odCostMatrix.FindOriginIndex(naLocation, curbApproach) : odCostMatrix.FindDestinationIndex(naLocation, curbApproach);
        }

        // Create the value that we are going to use as row and column headers in the output CSV file.
        // Currently, it takes the Name value from the input tables. Names can be duplicates, however,
        // so if there is a need for guaranteed uniqueness, a concatenation of OID and Name would suffice.
        // Replace any commas (and leading and training space with a single space to keep CSV structure 
        // correct.
        //
        var name = NormalizeCsvName((string)row.Value[nameFieldIndex]);
        if (!mapping.ContainsKey(internalMatrixIndex))
        {
          var originList = new List<string> { name };
          mapping.Add(internalMatrixIndex, originList);
        }
        else
        {
          mapping[internalMatrixIndex].Add(name);
        }
      }

      return mapping;
    }
  }
}