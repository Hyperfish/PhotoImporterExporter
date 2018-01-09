using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using Hyperfish.ImportExport.AD;
using Hyperfish.ImportExport.O365;
using Newtonsoft.Json;
using static Hyperfish.ImportExport.Log;

namespace Hyperfish.ImportExport
{
    class Program
    {
        private const string ReadPrompt = "> ";
        private const string ManifestFilename = "UserList.json";

        private static Dictionary<string, ImportExportRecord> _peopleList;

        static void Main(string[] args)
        {
            MainMenu();
        }
        
        static void MainMenu()
        {
            do
            {
                Console.Clear();

                WriteToConsole($"Hyperfish Export/Export tool" + Environment.NewLine);

                WriteToConsole($"About:");
                WriteToConsole($"This tool exports user profile photos from SharePoint Online and Exchange Online {Environment.NewLine}and can import photos to Active Directory. Photos are saved in /photos." + Environment.NewLine);

                WriteToConsole($"Options:");
                WriteToConsole($"1) Export photos from SharePoint Online");
                WriteToConsole($"2) Export photos from Exchange Online");
                WriteToConsole($"3) Import photos into Active Directory");
                WriteToConsole($"4) Import photos into SharePoint Online");
                WriteToConsole($"5) Import photos into Exchange Online");
                WriteToConsole($"6) Test import photos into Active Directory for one user");

                WriteToConsole($"7) Exit" + Environment.NewLine);
                
                var consoleInput = ReadFromConsole();
                if (string.IsNullOrWhiteSpace(consoleInput)) continue;

                try
                {
                    int option;

                    if (Int32.TryParse(consoleInput, out option))
                    {
                        switch (option)
                        {
                            // import
                            case 1:
                                ExportFromO365(O365Service.Spo);
                                break;

                            case 2:
                                ExportFromO365(O365Service.Exo);
                                break;

                            // export
                            case 3:
                                ImportToActiveDirectory(false);
                                break;

                            case 4:
                                ImportToO365(O365Service.Spo);
                                break;

                            case 5:
                                ImportToO365(O365Service.Exo);
                                break;

                            case 6:
                                ImportToActiveDirectory(true);
                                break;

                            case 7:
                                WriteToConsole("Exiting ... ");
                                return;
                        }
                    }
                }
                catch (Exception ex)
                {
                    // OOPS! Something went wrong - Write out the problem:
                    WriteToConsole(ex.Message);
                    WriteToConsole("Press any key to continue");
                    var input = ReadFromConsole();
                }

            } while (true);
            
        }
        
        static void ExportFromO365(O365Service service)
        {
            try
            {
                Logger.Debug($"Starting Export from {service}");

                WriteToConsole($"O365 Tenant name [e.g. \"Contoso\", from {{contoso}}.onmicrosoft.com]: ");
                var tenantName = ReadFromConsole();

                WriteToConsole($"O365 Administrator username [e.g. admin@contoso.onmicrosoft.com]: ");
                var username = ReadFromConsole();

                WriteToConsole($"O365 Administrator password: "); 
                var password = ReadSensitiveFromConsole();

                Logger.Debug($"Tenant: {tenantName}");
                Logger.Debug($"Username: {username}");

                var settings = new O365Settings()
                {
                    Username = username,
                    Password = password,
                    TenantName = tenantName
                };

                Logger.Debug($"Initializing O365 worker");

                var worker = new O365Worker(settings);

                var newPeopleList = worker.Export(service);

                // merge
                MergeAndSavePeopleLists(newPeopleList);
                
                Logger.Debug($"Completed export from {service}");

                WriteToConsole();
                WriteToConsole($"Completed export of photos. Press any key to return to main menu.");
                ReadFromConsole();

            }
            catch (Exception e)
            {
                WriteToConsole($"Problem exporting: {e.Message}");
                WriteToConsole("Press any key to continue");
                var input = ReadFromConsole();
            }
            
        }

        private static void MergeAndSavePeopleLists(Dictionary<string, ImportExportRecord> list)
        {
            foreach (var person in list)
            {
                if (!PeopleList.ContainsKey(person.Value.Upn))
                {
                    // add them
                    PeopleList[person.Value.Upn] = person.Value;
                }
                else
                {
                    // else update their photo location
                    PeopleList[person.Value.Upn].PhotoLocation = person.Value.PhotoLocation;
                }
            }

            SaveManifest(PeopleList);
        }

        private static void ImportToActiveDirectory(bool promptForUsersToImport)
        {
            try
            {
                Logger.Debug($"Starting import to AD");

                var testUser = string.Empty;

                if (promptForUsersToImport)
                {
                    WriteToConsole($"Upn for test user to import [e.g. aaron@contoso.com]: ");
                    testUser = ReadFromConsole();
                }

                WriteToConsole($"Active Directory domain contoller [e.g. 192.168.1.10, or domaincontroller.contoso.com]: ");
                var dc = ReadFromConsole();

                WriteToConsole($"Administrator username [e.g. administrator]: ");
                var username = ReadFromConsole();

                WriteToConsole($"Administrator password: ");
                var password = ReadSensitiveFromConsole();

                WriteToConsole($"Account domain [e.g. CORP]: ");
                var domain = ReadFromConsole();

                Logger.Debug($"Domain controller: {dc}");
                Logger.Debug($"Username: {username}");
                Logger.Debug($"Domain: {domain}");
                
                var settings = new AdSettings()
                {
                    AdServer = dc,
                    AdCredentials = new AdConnectionCredentials(username, password, domain)
                };

                Logger.Debug($"Initializing AD worker");

                var worker = new AdWorker(settings);

                if (promptForUsersToImport)
                {
                    if (!PeopleList.ContainsKey(testUser))
                    {
                        WriteToConsole($"Test user specified not found in user list.");
                    }
                    else
                    {
                        // make a small list
                        var testList = new Dictionary<string, ImportExportRecord>();
                        testList[testUser] = PeopleList[testUser];

                        // just import the test list
                        worker.ImportPicsToAd(testList);
                    }
                }
                else
                {
                    worker.ImportPicsToAd(PeopleList);
                }

                WriteToConsole();
                WriteToConsole($"Completed import of photos to AD. Press any key to return to main menu.");
                ReadFromConsole();

            }
            catch (Exception e)
            {
                WriteToConsole($"Problem importing: {e.Message}");
                WriteToConsole("Press any key to continue");
                var input = ReadFromConsole();
            }
        }

        private static void ImportToO365(O365Service service)
        {
            try
            {
                Logger.Debug($"Starting Import to from {service}");

                WriteToConsole($"O365 Tenant name [e.g. \"Contoso\", from {{contoso}}.onmicrosoft.com]: ");
                var tenantName = ReadFromConsole();

                WriteToConsole($"O365 Administrator username [e.g. admin@contoso.onmicrosoft.com]: ");
                var username = ReadFromConsole();

                WriteToConsole($"O365 Administrator password: ");
                var password = ReadSensitiveFromConsole();

                Logger.Debug($"Tenant: {tenantName}");
                Logger.Debug($"Username: {username}");

                var settings = new O365Settings()
                {
                    Username = username,
                    Password = password,
                    TenantName = tenantName
                };

                Logger.Debug($"Initializing O365 worker");

                var worker = new O365Worker(settings);

                worker.Import(PeopleList, service);
                
                Logger.Debug($"Completed import to {service}");

                WriteToConsole();
                WriteToConsole($"Completed import of photos. Press any key to return to main menu.");
                ReadFromConsole();

            }
            catch (Exception e)
            {
                WriteToConsole($"Problem importing: {e.Message}");
                WriteToConsole("Press any key to continue");
                var input = ReadFromConsole();
            }
        }
        
        private static string ReadFromConsole(string promptMessage = "")
        {
            // Show a prompt, and get input:
            Console.Write(ReadPrompt + promptMessage);
            return Console.ReadLine();
        }

        private static SecureString ReadSensitiveFromConsole(string promptMessage = "")
        {
            SecureString text = new SecureString();
            Console.Write(ReadPrompt + promptMessage);
            ConsoleKeyInfo key;

            do
            {
                key = Console.ReadKey(true);

                // Backspace Should Not Work
                if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
                {
                    text.AppendChar(key.KeyChar);
                    Console.Write("*");
                }
                else
                {
                    Console.Write("\b");
                }
            }
            // Stops Receving Keys Once Enter is Pressed
            while (key.Key != ConsoleKey.Enter);

            Console.WriteLine();

            return text;
        }

        private static void WriteToConsole(string message = "")
        {
            if (message.Length > 0)
            {
                Console.WriteLine(message);
            }
        }
        
        private static Dictionary<string, ImportExportRecord> LoadManifest()
        {
            if (File.Exists(ManifestFilename))
            {
                var list = JsonConvert.DeserializeObject<Dictionary<string, ImportExportRecord>>(File.ReadAllText(ManifestFilename));
                return list;
            }
            else
            {
                var list = new Dictionary<string, ImportExportRecord>();
                
                SaveManifest(list);

                return list;
            }
        }

        private static void SaveManifest(Dictionary<string, ImportExportRecord> userList)
        {
            var json = JsonConvert.SerializeObject(userList, Formatting.Indented);

            File.WriteAllText(ManifestFilename, json, Encoding.UTF8);
        }
        
        private static Dictionary<string, ImportExportRecord> PeopleList
        {
            get
            {
                if(_peopleList == null)
                {
                    _peopleList = LoadManifest();
                }

                return _peopleList;
            }
        }
    }
}
