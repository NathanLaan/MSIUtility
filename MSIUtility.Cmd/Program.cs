using System;
using System.IO;
using WindowsInstaller;

namespace MSIUtility.Cmd {

    class Program {

        //
        // TODO: Include the following command-line arguments: 
        //          -d                      : Delete the input file.
        //          -f "{pn}-v{pv}.msi"     : Specify the format of the output filename.
        //
        static void Main(string[] args) {
            string inputFileName = string.Empty;
            string outputPath = string.Empty;
            string productName = "[ProductName]";
            string productVersion;

            if (args.Length < 2) {
                Console.WriteLine("Missing input file name and output path parameters.");
                inputFileName = Console.ReadLine();
            }
            else {
                inputFileName = args[0];
                outputPath = args[1];
            }

            try {
                if (inputFileName.EndsWith(".msi", StringComparison.OrdinalIgnoreCase)) {
                    productName = GetMsiProperty(inputFileName, "ProductName");
                    productVersion = GetMsiProperty(inputFileName, "ProductVersion");
                }
                else {
                    return;
                }

                //
                // TODO: Allow user to specify the format string, something like "{pn}-v{pv}.msi"
                //
                string outputFileName = string.Format(outputPath + Path.DirectorySeparatorChar + "{0}-v{1}.msi", productName, productVersion);

                Console.WriteLine("Output File Name: " + outputFileName);
                File.Copy(inputFileName, outputFileName, true);

                //
                // TODO: Allow user to specify whether or not to delete the input file
                //
                //File.Delete(inputFileName);
            }
            catch (Exception ex) {
                Console.WriteLine(ex.Message);
            }
        }

        static string GetMsiProperty(string msiFile, string property) {
            string propertyValue = string.Empty;

            // Create an Installer instance  
            Type classType = Type.GetTypeFromProgID("WindowsInstaller.Installer");
            Object installerObj = Activator.CreateInstance(classType);
            Installer installer = installerObj as Installer;

            // Open the msi file for reading  
            // 0 - Read, 1 - Read/Write  
            Database database = installer.OpenDatabase(msiFile, 0);

            // retrieve the requested property
            string sql = String.Format("SELECT Value FROM Property WHERE Property='{0}'", property);
            View view = database.OpenView(sql);
            view.Execute(null);

            // Read in the fetched record  
            Record record = view.Fetch();
            if (record != null) {
                propertyValue = record.get_StringData(1);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(record);
            }
            view.Close();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(view);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(database);

            return propertyValue;
        }

    }

}