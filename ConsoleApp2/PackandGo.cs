using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Compression;
using System.Diagnostics;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;

namespace PackandGo.cs
{
    class Program
    {
        static void Main()
        {
            ModelDoc2 swModelDoc = default(ModelDoc2);
            ModelDocExtension swModelDocExt = default(ModelDocExtension);
            PackAndGo swPackAndGo = default(PackAndGo);
            SldWorks swApp = new SldWorks();

            /* To be used to more easily choose which assemblies to pack and go
            List<string> trNames = new List<string>(new string[] { "500", "525"});
            List<string> cmtNames = new List<string>(new string[] { "500", "525" });

            for (int i = 0; i <= trNames.Count - 1; i++)
            {
                trNames[i] = "TR-34-20-" + trNames[i];
            }
            for (int i = 0; i <= cmtNames.Count - 1; i++)
            {
                cmtNames[i] = "CMT-34-20-" + cmtNames[i];
            }
            */

            List<string> modelNames = new List<string>(new string[] { "TR-34-20-400", "TR-34-20-425", "TR-34-20-500", "TR-34-20-525", "TR-34-20-530", "TR-34-20-600", "TR-34-20-625", "TR-34-20-630", "TR-34-20-700", "TR-34-20-725", "TR-34-20-730", "TR-34-20-900", "TR-34-20-925", "TR-34-20-930",
                                                                      "CMT-34-20-400", "CMT-34-20-425", "CMT-34-20-500", "CMT-34-20-525", "CMT-34-20-530", "CMT-34-20-600", "CMT-34-20-625", "CMT-34-20-630", "CMT-34-20-700", "CMT-34-20-725", "CMT-34-20-730", "CMT-34-20-900", "CMT-34-20-925", "CMT-34-20-930",});
            int modelCount = modelNames.Count;


            for (int j = 0; j <= modelCount - 1; j++)
            {
                string openFile = null;
                bool status = false;
                int warnings = 0;
                int errors = 0;
                int i = 0;
                int namesCount = 0;
                string savePath = null;
                int[] statuses = null;

                openFile = @"C:\Configurator\" +modelNames[j] + ".sldasm";
                Debug.Print("Performing pack and go on " + modelNames[j]);
                swModelDoc = swApp.OpenDoc6(openFile, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
                swModelDocExt = (ModelDocExtension)swModelDoc.Extension;

                swPackAndGo = (PackAndGo)swModelDocExt.GetPackAndGo();

                savePath = @"C:\Configurator\PackandGoTest\test.zip";
                status = swPackAndGo.SetSaveToName(true, savePath);
                swPackAndGo.FlattenToSingleFolder = true;
                
                namesCount = swPackAndGo.GetDocumentNamesCount();
                Debug.Print("  Number of model documents: " + namesCount);

                // Include any drawings, SOLIDWORKS Simulation results, and SOLIDWORKS Toolbox components
                swPackAndGo.IncludeDrawings = true;
                Debug.Print(" Include drawings: " + swPackAndGo.IncludeDrawings);
                swPackAndGo.IncludeSimulationResults = true;
                Debug.Print(" Include SOLIDWORKS Simulation results: " + swPackAndGo.IncludeSimulationResults);
                swPackAndGo.IncludeToolboxComponents = true;
                Debug.Print(" Include SOLIDWORKS Toolbox components: " + swPackAndGo.IncludeToolboxComponents);

                // Verify document paths and filenames after adding prefix and suffix
                object getFileNames;
                object getDocumentStatus;
                string[] pgGetFileNames = new string[namesCount - 1];

                status = swPackAndGo.GetDocumentSaveToNames(out getFileNames, out getDocumentStatus);
                pgGetFileNames = (string[])getFileNames;
                /*This section is unnecessary and clutters the debug window, add back in at own preference
                Debug.Print("");
                Debug.Print("  My Pack and Go path and filenames after adding prefix and suffix: ");
                for (i = 0; i <= namesCount - 1; i++)
                {
                    Debug.Print("    My path and filename is: " + pgGetFileNames[i]);
                }
                */

                // Pack and Go
                statuses = (int[])swModelDocExt.SavePackAndGo(swPackAndGo);
                swApp.CloseDoc(modelNames[j]);
            }
        }
    }
}
