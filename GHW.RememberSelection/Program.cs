using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace GHW.RememberSelection
{
    using TSM = Tekla.Structures.Model;

    class Program
    {
        static void Main(string[] args)
        {
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

            Execute(tempFolder, ncTemplate, 1);
        }

        private static void Execute(in string tempFolder, in string ncTemplate, int attempt)
        {
            if (!Directory.Exists(tempFolder))
            {
                // Necessary otherwise the NC output fails.
                Directory.CreateDirectory(tempFolder);
            }

            Console.WriteLine($"Attempt no #{attempt}.");

            var modelObjectSelector = new TSM.UI.ModelObjectSelector();
            var selectedObjects = modelObjectSelector.GetSelectedObjects();
            selectedObjects.SelectInstances = false;

            // Create sets of unique parts.
            var allParts = new List<KeyValuePair<TSM.Part, string>>();
            var uniqueCombinations = new List<Dictionary<string, TSM.Part>>();

            while (selectedObjects.MoveNext())
            {
                var modelObject = selectedObjects.Current;

                if (modelObject is TSM.Part part)
                {
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
                var modelObjects = new ArrayList();
                foreach (var modelObject in allParts)
                {
                    modelObjects.Add(modelObject.Key);
                }

                var select = true;
                if (!modelObjectSelector.Select(modelObjects))
                {
                    select = false;
                }

                if (!select)
                {
                    Console.WriteLine("Select failed");
                }

                int size = modelObjectSelector.GetSelectedObjects().GetSize();

                if (size != allParts.Count)
                {
                    // This is wrong... The initial selection is not selected again.
                    Console.WriteLine($"Not all objects selected, {size}/{allParts.Count} selected.");
                    Debugger.Break();
                    Console.ReadKey();
                }
                else
                {
                    Execute(tempFolder, ncTemplate, ++attempt);
                }
            }
        }
    }
}
