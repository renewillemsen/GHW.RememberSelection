using CsvHelper;
using CsvHelper.Configuration;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;

namespace GHW.RememberSelection
{
    using TSM = Tekla.Structures.Model;

    class Program
    {
        static void Main(string[] args)
        {
            string logFile = $@"C:\temp\missing-parts-{DateTime.Now.ToString("yyyyMMddHHmmss")}.csv";

            Console.WriteLine($"TSM.Model assembly version: {typeof(TSM.Model).Assembly.GetName().Version}");

            string tempFolder = args != null && args.Length >= 1 ? args[0] : null; // @"C:\temp\NC_Files\";
            string ncTemplate = args != null && args.Length >= 2 ? args[1] : null; // "Alles";

            if (string.IsNullOrEmpty(tempFolder))
            {
                Console.Write("Please specify a temporary folder: ");
                tempFolder = Console.ReadLine();
            }

            if (string.IsNullOrEmpty(ncTemplate))
            {
                Console.Write("Please specify a NC Template: ");
                ncTemplate = Console.ReadLine();
            }

            Console.WriteLine(string.Empty);

            Execute(tempFolder, ncTemplate, 1, logFile);
        }

        private static void Execute(in string tempFolder, in string ncTemplate, int attempt, in string logFile)
        {
            if (!Directory.Exists(tempFolder))
            {
                // Necessary otherwise the NC output fails.
                Directory.CreateDirectory(tempFolder);
            }

            Console.WriteLine($"Attempt no #{attempt}.");

            var modelObjectSelector = new TSM.UI.ModelObjectSelector();
            var selectedObjects = modelObjectSelector.GetSelectedObjects();

            // Create sets of unique parts.
            var allParts = new List<KeyValuePair<TSM.Part, string>>();
            var uniqueCombinations = new List<Dictionary<string, TSM.Part>>();
            var allObjects = new ArrayList();
            while (selectedObjects.MoveNext())
            {
                var modelObject = selectedObjects.Current;

                if (modelObject is TSM.Part part)
                {
                    allObjects.Add(modelObject);
                    var partNo = part.GetPartMark();
                    allParts.Add(new KeyValuePair<TSM.Part, string>(part, partNo));

                    var found = false;

                    foreach (var dict in uniqueCombinations)
                    {
                        if (!dict.ContainsKey(partNo))
                        {
                            dict.Add(partNo, part);
                            found = true;
                            break;
                        }
                    }

                    if (!found)
                    {
                        var newDictionary = new Dictionary<string, TSM.Part>();
                        newDictionary.Add(partNo, part);
                        uniqueCombinations.Add(newDictionary);
                    }
                }
            }

            // Create NC files for each unique combination.
            try
            {
                var modelObjects = new ArrayList();

                foreach (var dict in uniqueCombinations)
                {
                    foreach (var part in dict.Values)
                    {
                        modelObjects.Add(part);
                    }

                    var select = modelObjectSelector.Select(modelObjects);

                    if (!select)
                    {
                        Console.WriteLine("Select failed");
                    }

                    // Remark: This method to create NC files sometimes gives a false positive: Returns true where there are no files created!
                    var createNCFiles = TSM.Operations.Operation.CreateNCFilesFromSelected(ncTemplate, tempFolder);
                    if (!createNCFiles)
                    {
                        Console.WriteLine("CreateNCFilesFromSelected failed.");
                    }

                    modelObjects.Clear();
                }
            }
            finally
            {
                // Reset the selection to the original state.
                var select = true;
                if (!modelObjectSelector.Select(allObjects))
                {
                    select = false;
                }

                if (!select)
                {
                    Console.WriteLine("Select failed");
                }

                var selector = modelObjectSelector.GetSelectedObjects();
                int size = selector.GetSize();

                if (size != allParts.Count)
                {
                    var missing = GetMissingParts(allParts.Select(t => t.Key), GetSelectedParts(selector));

                    WriteToCsv(missing, logFile);

                    // This is wrong... The initial selection is not selected again.
                    Console.WriteLine($"Not all objects selected, {size}/{allParts.Count} selected.");
                    Debugger.Break();
                    Console.ReadKey();
                }
                else
                {
                    Execute(tempFolder, ncTemplate, ++attempt, logFile);
                }
            }
        }

        private static void WriteToCsv(in IEnumerable<TSM.Part> missing, in string logFile)
        {
            using (var writer = new StreamWriter(logFile))
            {
                var csvConfig = new CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    Delimiter = ";"
                };

                using (var csv = new CsvWriter(writer, csvConfig))
                {
                    foreach (var part in missing)
                    {
                        dynamic t = new
                        {
                            guid = part.Identifier.GUID,
                            id = part.Identifier.ID,
                            id2 = part.Identifier.ID2,
                            name = part.GetPartMark(),
                            profile = part.Profile.ProfileString,
                            isUpToDate = part.IsUpToDate,
                            typeOf = part.GetType().FullName
                        };

                        csv.WriteRecord(t);
                        csv.NextRecord();
                    }
                }
            }
        }

        private static IEnumerable<TSM.Part> GetMissingParts(in IEnumerable<TSM.Part> allParts, in IEnumerable<TSM.Part> selectedParts)
        {
            var missing = new List<TSM.Part>();

            foreach (var part in allParts)
            {
                var found = false;
                foreach (var selected in selectedParts)
                {
                    if (selected.Identifier.Equals(part.Identifier))
                    {
                        found = true;
                        break;
                    }
                }

                if (!found)
                {
                    missing.Add(part);
                }
            }

            return missing;
        }

        private static IEnumerable<TSM.Part> GetSelectedParts(in TSM.ModelObjectEnumerator selectedParts)
        {
            var result = new List<TSM.Part>();

            while (selectedParts.MoveNext())
            {
                var part = selectedParts.Current as TSM.Part;

                if (part != null)
                {
                    result.Add(part);
                }
            }

            return result;
        }
    }
}
