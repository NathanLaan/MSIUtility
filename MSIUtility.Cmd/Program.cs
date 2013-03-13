using System;
using System.IO;
using WindowsInstaller;

namespace MSIUtility.Cmd
{
    class Program
    {
        static void Main(string[] args)
        {
            string inputFileName;
            string productName = "[ProductName]";
            string productVersion;

            if (args.Length == 0)
            {
                Console.WriteLine("Enter MSI filename: ");
                inputFileName = Console.ReadLine();
            }
            else
            {
                inputFileName = args[0];
            }

            try
            {

                if (inputFileName.EndsWith(".msi", StringComparison.OrdinalIgnoreCase))
                {
                    productName = GetMsiProperty(inputFileName, "ProductName");
                    productVersion = GetMsiProperty(inputFileName, "ProductVersion");
                }
                else
                {
                    return;
                }

                File.Copy(inputFileName, string.Format("{0}-{1}.msi", productName, productVersion));
                File.Delete(inputFileName);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        static string GetMsiProperty(string msiFile, string property)
        {
            string retVal = string.Empty;

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
            if (record != null)
            {
                retVal = record.get_StringData(1);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(record);
            }
            view.Close();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(view);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(database);

            return retVal;
        }
    }
}
