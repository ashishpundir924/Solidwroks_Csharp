using SolidWorks.Interop.sldworks;
using System;
using System.IO;

namespace CreateStepFiles
{
    class Program
    {
        static SldWorks swApp;

        static void Main(string[] args)
        {
            string directoryName = GetDirectoryName();

            if (!GetSolidWorks())
            {
                return;
            }

            int i = 0;

            foreach (string fileName in Directory.GetFiles(directoryName))
            {
                if (Path.GetExtension(fileName).ToLower() == ".sldprt")
                {
                    CreateStepFile(fileName, 1);
                    i += 1;
                }
                else if (Path.GetExtension(fileName).ToLower() == ".sldasm")
                {
                    CreateStepFile(fileName, 2);
                    i += 1;
                }
            }

            Console.WriteLine("Finished converting {0} files", i);

        }

        static void CreateStepFile(string fileName, int docType)
        {
            int errors = 0;
            int warnings = 0;

            ModelDoc2 swModel = swApp.OpenDoc6(fileName, docType, 1, "", ref errors, ref warnings);

            string stepFile = Path.Combine(Path.GetDirectoryName(fileName), Path.GetFileNameWithoutExtension(fileName), ".STEP");

            swModel.Extension.SaveAs(stepFile, 0, 1, null, ref errors, ref warnings);

            Console.WriteLine("Created STEP file: " + stepFile);;

            swApp.CloseDoc(fileName);
        }

        static string GetDirectoryName()
        {
            Console.WriteLine("Directory to Converty");
            string s = Console.ReadLine();

            if (Directory.Exists(s))
            {
                return s;
            }

            Console.WriteLine("Directory does not exists, try again");
            return GetDirectoryName();
        }

        static bool GetSolidWorks()
        {
            try
            {
                swApp = (SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"));

                if (swApp == null)
                {
                    throw new NullReferenceException(nameof(swApp));
                }

                if (!swApp.Visible)
                {
                    swApp.Visible = true;
                }

                Console.WriteLine("SolidWorks Loaded");
                return true;
            }
            catch (Exception)
            {
                Console.WriteLine("Could not launch SolidWorks");
                return false;
            }
        }
    }
}
