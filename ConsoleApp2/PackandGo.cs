using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Text;
using System.Threading.Tasks;
using System.IO.Compression;
using System.Diagnostics;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Runtime.InteropServices.ComTypes;


namespace PackandGo.cs
{
    class Program
    {
        static void Main()
        {
            ModelDoc2 swModelDoc = default(ModelDoc2);
            ModelDocExtension swModelDocExt = default(ModelDocExtension);
            PackAndGo swPackAndGo = default(PackAndGo);
            //SldWorks swApp = new SldWorks();
            Tuple<SldWorks, Process> processes = GetSolidworks.Solidworks(1);
            SldWorks swApp = processes.Item1;
            //add another line to get item 2 which is the process which you can then use to kill when it derps
            
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

            //List<string> modelNames = new List<string>(new string[] { "TR-34-20-400", "TR-34-20-425", "TR-34-20-500", "TR-34-20-525", "TR-34-20-530", "TR-34-20-600", "TR-34-20-625", "TR-34-20-630", "TR-34-20-700", "TR-34-20-725", "TR-34-20-730", "TR-34-20-900", "TR-34-20-925", "TR-34-20-930",
            //                                                          "CMT-34-20-400", "CMT-34-20-425", "CMT-34-20-500", "CMT-34-20-525", "CMT-34-20-530", "CMT-34-20-600", "CMT-34-20-625", "CMT-34-20-630", "CMT-34-20-700", "CMT-34-20-725", "CMT-34-20-730", "CMT-34-20-900", "CMT-34-20-925", "CMT-34-20-930",});
            List<string> modelNames = new List<string>(new string[] { 
                                                                      "CMT-34-20-730"});
            int modelCount = modelNames.Count;


            for (int j = 0; j <= modelCount - 1; j++)
            {
                try
                {
                    string openFile = null;
                    bool status = false;
                    int warnings = 0;
                    int errors = 0;
                    int i = 0;
                    int namesCount = 0;
                    string savePath = null;
                    int[] statuses = null;

                    openFile = @"C:\Configurator\" + modelNames[j] + ".sldasm";
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
                catch
                {
                    /*
                    Process swProcess = new Process();
                    int processID = swApp.GetProcessID();
                    swProcess = Process.GetProcessById(processID);
                    swProcess.Kill();
                    while (swProcess.HasExited == false) { }
                    */
                    swApp = null;
                    Process Solidworks = null;
                    Debug.Print("Solidworks Crash: Restarting Now");
                        while (swApp == null)
                        {
                            Tuple<SldWorks, Process> tuple = GetSolidworks.Solidworks(2);
                            swApp = tuple.Item1;
                            Solidworks = tuple.Item2;

                        }
                }
            }
        }
    }
}

static class GetSolidworks
{
    private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
    static public Tuple<SldWorks, Process> Solidworks(int Num)
    {
        System.Diagnostics.Process SolidWorksPrc = null;
        if (SolidWorksPath() != null)
        {
            RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\SolidWorks\SOLIDWORKS 2018\ExtReferences", true);
            key.SetValue("SolidWorks Journal Folders", @"C:\Users\tsowers\Documents\Journals\" + Num.ToString());
            SolidWorksPrc = System.Diagnostics.Process.Start(SolidWorksPath());

            System.Threading.Thread.Sleep(25000);
            key.SetValue("SolidWorks Journal Folders", @"C:\Users\tsowers\AppData\Roaming\SolidWorks\SOLIDWORKS 2018");
            key.Dispose();



        }



        if (SolidWorksPrc == null)
        {
            Debug.Print(Num.ToString() + " failed to start");
            //return null;
            Tuple<SldWorks, Process> tuple = new Tuple<SldWorks, Process>(null, SolidWorksPrc);
            return tuple;
        }


        int ID = SolidWorksPrc.Id;

        try
        {
            SldWorks sldWorks = (SldWorks)ROTHelper.GetActiveObjectList(ID.ToString()).Where(keyvalue => (keyvalue.Key.ToLower().Contains("solidworks")))
            .Select(keyvalue => keyvalue.Value)
            .First();
            Tuple<SldWorks, Process> tuple = new Tuple<SldWorks, Process>(sldWorks, SolidWorksPrc);
            return tuple;
        }
        catch (Exception ex)
        {
            logger.Error(ex, "Error marshalling the solidworks object");
            Debug.Print(Num.ToString() + " Failed to get SldWorks object, will try again");
            //Thread.Sleep(300);
            SolidWorksPrc.Kill();
            Tuple<SldWorks, Process> tuple = new Tuple<SldWorks, Process>(null, SolidWorksPrc);
            return tuple;
            //SolidWorks(Num);
        }

        /*
        try
        {
            return (SldWorks, SolidWorksPrc)
                ROTHelper.GetActiveObjectList(ID.ToString())
            .Where(
                keyvalue => (keyvalue.Key.ToLower().Contains("solidworks")))
            .Select(keyvalue => keyvalue.Value)
            .First();
        }
        catch
        {
            Debug.Print(Num.ToString() + " Failed to get SldWorks object, will try again");
            Thread.Sleep(300);
            try
            {
                return (SldWorks)
                ROTHelper.GetActiveObjectList(ID.ToString())
            .Where(
                keyvalue => (keyvalue.Key.ToLower().Contains("solidworks")))
            .Select(keyvalue => keyvalue.Value)
            .First();
            }
            catch
            {
                Debug.Print(Num.ToString() + " Failed on second try to get solidworks object");
                SolidWorksPrc.Kill();
                return null;
            }

        }
        */
    }

    static private string SolidWorksPath()
    {
        return @"C:\Program Files\SolidWorks Corp\SolidWorks\SLDWORKS.exe";
        // More complicated in actual program - takes a look at selected registry locations to guess where solidworks is.  
    }
}


/// <summary>  
/// The COM running object table utility class.  
/// </summary>  
class ROTHelper
{
    #region APIs  

    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

    [DllImport("ole32.dll", PreserveSig = false)]
    private static extern void CLSIDFromProgIDEx([MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid clsid);

    [DllImport("ole32.dll", PreserveSig = false)]
    private static extern void CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid clsid);

    [DllImport("ole32.dll")]
    private static extern int ProgIDFromCLSID([In()]ref Guid clsid, [MarshalAs(UnmanagedType.LPWStr)]out string lplpszProgID);

    #endregion

    #region Public Methods  

    /// <summary>  
    /// Converts a COM class ID into a prog id.  
    /// </summary>  
    /// <param name="progID">The prog id to convert to a class id.</param>  
    /// <returns>Returns the matching class id or the prog id if it wasn't found.</returns>  
    public static string ConvertProgIdToClassId(string progID)
    {
        Guid testGuid;
        try
        {
            CLSIDFromProgIDEx(progID, out testGuid);
        }
        catch
        {
            try
            {
                CLSIDFromProgID(progID, out testGuid);
            }
            catch
            {
                return progID;
            }
        }
        return testGuid.ToString().ToUpper();
    }

    /// <summary>  
    /// Converts a COM class ID into a prog id.  
    /// </summary>  
    /// <param name="classID">The class id to convert to a prog id.</param>  
    /// <returns>Returns the matching class id or null if it wasn't found.</returns>  
    public static string ConvertClassIdToProgId(string classID)
    {
        Guid testGuid = new Guid(classID.Replace("!", ""));
        string progId = null;
        try
        {
            ProgIDFromCLSID(ref testGuid, out progId);
        }
        catch (Exception)
        {
            return null;
        }
        return progId;
    }

    /// <summary>  
    /// Get a snapshot of the running object table (ROT).  
    /// </summary>  
    /// <param name="filter">The filter to apply to the list (nullable).</param>  
    /// <returns>A hashtable of the matching entries in the ROT</returns>  
    public static Dictionary<string, object> GetActiveObjectList(string filter)
    {
        Dictionary<string, object> result = new Dictionary<string, object>();

        IntPtr numFetched = new IntPtr();
        IRunningObjectTable runningObjectTable;
        IEnumMoniker monikerEnumerator;
        IMoniker[] monikers = new IMoniker[1];

        IBindCtx ctx;
        CreateBindCtx(0, out ctx);

        ctx.GetRunningObjectTable(out runningObjectTable);
        runningObjectTable.EnumRunning(out monikerEnumerator);
        monikerEnumerator.Reset();

        while (monikerEnumerator.Next(1, monikers, numFetched) == 0)
        {
            string runningObjectName;
            monikers[0].GetDisplayName(ctx, null, out runningObjectName);

            if (filter == null || filter.Length == 0 || runningObjectName.IndexOf(filter) != -1)
            {
                object runningObjectVal;
                runningObjectTable.GetObject(monikers[0], out runningObjectVal);
                result[runningObjectName] = runningObjectVal;
            }
        }

        return result;
    }

    /// <summary>  
    /// Returns an object from the ROT, given a prog Id.  
    /// </summary>  
    /// <param name="progId">The prog id of the object to return.</param>  
    /// <returns>The requested object, or null if the object is not found.</returns>  
    public static object GetActiveObject(string progId)
    {
        // Convert the prog id into a class id  
        string classId = ConvertProgIdToClassId(progId);

        IRunningObjectTable runningObjectTable = null;
        IEnumMoniker monikerEnumerator = null;
        IBindCtx ctx = null;
        try
        {
            IntPtr numFetched = new IntPtr();
            // Open the running objects table.  
            CreateBindCtx(0, out ctx);
            ctx.GetRunningObjectTable(out runningObjectTable);
            runningObjectTable.EnumRunning(out monikerEnumerator);
            monikerEnumerator.Reset();
            IMoniker[] monikers = new IMoniker[1];

            // Iterate through the results  
            while (monikerEnumerator.Next(1, monikers, numFetched) == 0)
            {

                string runningObjectName;
                monikers[0].GetDisplayName(ctx, null, out runningObjectName);
                if (runningObjectName.IndexOf(classId) != -1)
                {
                    // Return the matching object  
                    object objReturnObject;
                    runningObjectTable.GetObject(monikers[0], out objReturnObject);
                    return objReturnObject;
                }
            }
            return null;
        }
        finally
        {
            // Free resources  
            if (runningObjectTable != null)
                Marshal.ReleaseComObject(runningObjectTable);
            if (monikerEnumerator != null)
                Marshal.ReleaseComObject(monikerEnumerator);
            if (ctx != null)
                Marshal.ReleaseComObject(ctx);
        }
    }

    #endregion
}
