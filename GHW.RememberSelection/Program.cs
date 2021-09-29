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
            const string tempFolder = @"C:\temp\NC_Files\";
            const string ncTemplate = "Alles"; // DSTV for Profiles

            if (!Directory.Exists(tempFolder))
            {
                Directory.CreateDirectory(tempFolder);
            }

            int attempt = args != null && args.Length == 1 ? Convert.ToInt32(args[0]) : 0;

            Console.WriteLine($"Attempt no #{attempt}.");

            var modelObjectSelector = new TSM.UI.ModelObjectSelector();
            var selectedObjects = modelObjectSelector.GetSelectedObjects();
            selectedObjects.SelectInstances = false;

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
                    modelObjects.Add(modelObject);
                }

                var select = modelObjectSelector.Select(modelObjects);
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
                    Main(new string[] { (++attempt).ToString() });
                }
            }
        }
    }
}
