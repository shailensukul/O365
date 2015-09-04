using O365.Documentor.Inventory.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace O365.Documentor.Inventory
{
    class Program
    {

        private static void TagLine()
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("\n-------------------------------");
            Console.WriteLine("Author: Shailen Sukul ( @shailensukul )");
            Console.WriteLine("SHAREPOINT CLOUD DESIGN");
            Console.WriteLine("http://wwww.SharePointCloudDesign.com");
            Console.WriteLine("-------------------------------");
            Console.ForegroundColor = ConsoleColor.Green;
        }

        private static void Header()
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.BackgroundColor = ConsoleColor.Black;

            Console.WriteLine("\n-------------------------------");
            Console.WriteLine("OFFICE 365 Documentor");
            Console.WriteLine("-------------------------------");
            Console.WriteLine("***** USE AT YOUR OWN RISK *****");
            Console.ForegroundColor = ConsoleColor.Green;
        }

        private static void Variables()
        {
            Console.ForegroundColor = ConsoleColor.Yellow;

            Console.WriteLine("\n-------------------------------");
            Console.WriteLine("CONFIGURATION");
            Console.WriteLine("-------------------------------");
            Console.WriteLine(string.Format("Output folder: {0}", ConfigurationManager.Instance.OutputDirectory));
            Console.WriteLine(string.Format("Site Collection URL: {0}", ConfigurationManager.Instance.SiteCollectionUrl));
            Console.WriteLine(string.Format("Content Type Group Filter: {0}", ConfigurationManager.Instance.ContentTypeGroupName));
            Console.WriteLine("Refer to the O365.Documentor.exe.config file to update config.");
            Console.WriteLine("-------------------------------");
            Console.ForegroundColor = ConsoleColor.Green;
        }

        static void Main(string[] args)
        {
            bool loop = true;
            while (loop)
            {
                try
                {
                    Console.Clear();

                    Header();

                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("\n----------------");
                    Console.WriteLine("Choose an option");
                    Console.WriteLine("----------------");
                    Console.WriteLine("1 - View configuration");
                    Console.WriteLine("2 - Site Inventory Report (CSV)");
                    Console.WriteLine("3 - Site Inventory Report (XML)");
                    Console.WriteLine("4 - Content Type Inventory Report");
                    Console.WriteLine("5 - Site Columns Inventory Report");
                    Console.WriteLine("6 - PageLayouts Inventory Report");
                    Console.WriteLine("\nENTER To Exit");
                    TagLine();
                    string key = Console.ReadLine();
                    switch (key)
                    {
                        case "1":
                            Console.Clear();
                            Variables();
                            Console.WriteLine("Press ENTER to return");
                            Console.ReadLine();
                            break;
                        case "2":
                            Console.Clear();
                            Console.WriteLine("Collating site inventory. Please be patient, this can take a long time...");
                            SiteCollectionInventory.Instance.Init();
                            SiteCollectionInventory.Instance.Execute(
                                ConfigurationManager.Instance.SiteCollectionUrl);
                            Console.WriteLine(
                                string.Format("Done. Check output directory ({0}).\nPress enter to return.",
                                ConfigurationManager.Instance.OutputDirectory
                                ));
                            Console.ReadLine();
                            break;
                        case "3":
                            Console.Clear();
                            Console.WriteLine("Collating site inventory. Please be patient, this can take a long time...");
                            SiteInventory.Instance.Init();
                            SiteInventory.Instance.Execute(
                                ConfigurationManager.Instance.SiteCollectionUrl);
                            Console.WriteLine(
                                string.Format("Done. Check output directory ({0}).\nPress enter to return.",
                                ConfigurationManager.Instance.OutputDirectory
                                ));
                            Console.ReadLine();
                            break;
                        case "4":
                            Console.Clear();
                            Console.WriteLine("Exporting content types..");
                            ContentTypeInventory.Instance.Init();
                            ContentTypeInventory.Instance.Execute(
                                ConfigurationManager.Instance.SiteCollectionUrl);
                            Console.WriteLine(
                                string.Format("Done. Check output directory ({0}).\nPress enter to return.",
                                ConfigurationManager.Instance.OutputDirectory
                                ));
                            Console.ReadLine();
                            break;
                        case "5":
                            Console.Clear();
                            Console.WriteLine("Exporting site columns..");
                            SiteColumnInventory.Instance.Init();
                            SiteColumnInventory.Instance.Execute(
                                ConfigurationManager.Instance.SiteCollectionUrl);
                            Console.WriteLine(
                                string.Format("Done. Check output directory ({0}).\nPress enter to return.",
                                ConfigurationManager.Instance.OutputDirectory
                                ));
                            Console.ReadLine();
                            break;
                        case "6":
                            Console.Clear();
                            Console.WriteLine("Exporting page layouts.");
                            PageLayoutInventory.Instance.Init();
                            PageLayoutInventory.Instance.Execute(
                                ConfigurationManager.Instance.SiteCollectionUrl);
                            Console.WriteLine(
                                string.Format("Done. Check output directory ({0}).\nPress enter to return.",
                                ConfigurationManager.Instance.OutputDirectory
                                ));
                            Console.ReadLine();
                            break;
                        case "":
                            loop = false;
                            break;
                        default:
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("Invalid option. Press enter to continue.");
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.ReadLine();
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    Console.WriteLine(ex.ToString());
                    Console.ForegroundColor = ConsoleColor.DarkGreen;
                    Console.WriteLine("Press enter to continue.");
                    Console.ReadLine();
                }

            }

        }
    }
}
