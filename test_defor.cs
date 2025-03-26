using System;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

class Program
{
    static void Main()
    {
        Console.WriteLine("Enter the full path to the folder containing .sldprt files:");
        string sourceFolder = Console.ReadLine();

        Console.WriteLine("Enter the full path to the folder to save .stp files:");
        string outputFolder = Console.ReadLine();

        if (!Directory.Exists(sourceFolder) || !Directory.Exists(outputFolder))
        {
            Console.WriteLine("Invalid folder path(s). Exiting.");
            return;
        }

        SldWorks swApp = Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")) as SldWorks;
        swApp.Visible = true;

        foreach (string file in Directory.GetFiles(sourceFolder, "*.sldprt"))
        {
            Console.WriteLine($"\nProcessing: {Path.GetFileName(file)}");
            ModelDoc2 model = swApp.OpenDoc6(file, (int)swDocumentTypes_e.swDocPART,
                                              (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", 
                                              out _, out _);

            if (model == null)
            {
                Console.WriteLine("Failed to open file.");
                continue;
            }

            ConfigurationManager configMgr = model.ConfigurationManager;
            Configuration[] configs = configMgr.GetConfigurations() as Configuration[];

            bool multipleConfigs = configs != null && configs.Length > 1;

            foreach (Configuration config in configs)
            {
                string configName = config.Name;
                configMgr.ActivateConfiguration(configName);

                string outputFile = Path.Combine(outputFolder,
                    $"{Path.GetFileNameWithoutExtension(file)}_{configName}.stp");

                int status = model.Extension.SaveAs(outputFile,
                    (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                    (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                    null, ref var _, ref var _);

                Console.WriteLine($"Exported: {outputFile}");
            }

            // Close document after export
            swApp.CloseDoc(Path.GetFileName(file));
        }

        Console.WriteLine("\nExport complete. Press any key to exit...");
        Console.ReadKey();
    }
}
