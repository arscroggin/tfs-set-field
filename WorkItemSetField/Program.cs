using System;
using Microsoft.TeamFoundation;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Framework.Client;
using Microsoft.TeamFoundation.Framework.Common;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace WorkItemSetField
{
    class Program
    {
        private static TfsTeamProjectCollection _tfs;
        private static int _workItemNumber;
        private static string _fieldName;
        private static string _fieldValue;


        // setfield -w <work item number> -f <field> -v <value>
        static void Main(string[] args)
        {
            ProcessCommandLine(args);
            ConnectToTfs();
            UpdateWorkItem();
        }

        private static void ProcessCommandLine(string[] args)
        {
            if (args.Length != 6)
            {
                PrintHelp();
                Environment.Exit(-1);
            }

            try
            {

                for (int i = 0; i < args.Length; i += 2)
                {
                    switch (args[i])
                    {
                        case "-w":
                            _workItemNumber = Convert.ToInt32(args[i + 1]);
                            break;

                        case "-f":
                            _fieldName = args[i + 1];
                            break;

                        case "-v":
                            _fieldValue = args[i + 1];
                            break;

                        default:
                            PrintHelp();
                            Environment.Exit(-1);
                            break;
                    }
                }
            }
            catch (Exception)
            {
                PrintHelp();
                Environment.Exit(-1);
            }
        }

        private static void ConnectToTfs()
        {
            var serverUri = new Uri("https://tfsclient.sep.com:8383/tfs/Collection02");
            Console.WriteLine("Connecting to " + serverUri);

            try
            {
                _tfs = new TfsTeamProjectCollection(serverUri);

                // Get the identity to impersonate
                var ims = _tfs.GetService<IIdentityManagementService>();

                var identity = ims.ReadIdentity(IdentitySearchFactor.AccountName, @"OC\tfs_user",
                   MembershipQuery.None, ReadIdentityOptions.None);

                _tfs = new TfsTeamProjectCollection(serverUri, identity.Descriptor);
                Console.WriteLine("Connected");
            }
            catch (TeamFoundationServiceUnavailableException)
            {
                Console.WriteLine("FAILED: TFS server not available");
                Environment.Exit(-1);
            }

        }

        private static void PrintHelp()
        {
            Console.WriteLine("usage: setfield");
            Console.WriteLine("   -w <work item number>");
            Console.WriteLine("   -f <field name>");
            Console.WriteLine("   -v \"value to set\"");
        }

        static void UpdateWorkItem()
        {
            var store = (WorkItemStore)_tfs.GetService(typeof(WorkItemStore));

            try
            {
                var workItem = store.GetWorkItem(_workItemNumber);
                workItem.Fields[_fieldName].Value = _fieldValue;
                Console.WriteLine("Work item updated");
                workItem.Save();
            }
            catch (DeniedOrNotExistException)
            {
                Console.WriteLine("Could not find work item " + _workItemNumber);
                Environment.Exit(-1);
            }
            catch (FieldDefinitionNotExistException)
            {
                Console.WriteLine("Could not find the field you were looking for...");
                Environment.Exit(-1);
            }
            catch (TeamFoundationServiceUnavailableException)
            {
                Console.WriteLine("FAILED: TFS server not available");
                Environment.Exit(-1);
            }

        }
    }
}
