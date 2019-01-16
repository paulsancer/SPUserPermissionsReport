using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;

namespace SP.UserPermissionsReport
{
    internal class App
    {
        private readonly string _tenantAdminUrl;
        private readonly string _clientId;
        private readonly string _clientSecret;
        private readonly ITenantService _svc;

        public App()
        {
            Console.WriteLine("# User Permissions Report Tool");
            Console.WriteLine("  Use this tool to get a list of SharePoint permissions for a specific user in one or multiple site collections\n\n");

            _tenantAdminUrl = ConfigurationManager.AppSettings["tenantAdminUrl"];
            _clientId = ConfigurationManager.AppSettings["clientId"];
            _clientSecret = ConfigurationManager.AppSettings["clientSecret"];

            if (string.IsNullOrEmpty(_tenantAdminUrl))
            {
                Console.Write("Enter the SharePoint tenant admin URL: ");
                _tenantAdminUrl = Console.ReadLine();
            }

            Console.WriteLine("\nConnecting to tenant '{0}'...", _tenantAdminUrl);

            try
            {
                if (!string.IsNullOrEmpty(_clientId) && !string.IsNullOrEmpty(_clientSecret))
                {
                    Console.WriteLine("\nGetting App Only context using Client ID & Secret from app settings...\n");
                    try
                    {
                        _svc = new TenantService(_tenantAdminUrl, _clientId, _clientSecret);
                        Console.WriteLine("Connected!\n");
                    }
                    catch (Exception e)
                    {
                        Console.Error.WriteLine("Error:\n");
                        Console.Error.WriteLine(e.ToString() + "\n");
                    }
                }

                if (_svc == null)
                {
                    Console.Error.WriteLine("\nUnable to get App Only context using Client ID & Secret from app settings...");
                    Console.Write("\nDo you want to connect using web login?\nTenant administrator credentials will be required [Y = Yes / N = No, close program]: ");
                    string k = Console.ReadLine();
                    if (k.Equals("Y", StringComparison.InvariantCultureIgnoreCase))
                    {
                        Console.WriteLine("\nWaiting for web login context...");
                        try
                        {
                            _svc = new TenantService(_tenantAdminUrl);
                            Console.WriteLine("\nConnected!");
                        }
                        catch (Exception e)
                        {
                            if (e.Message.Contains("(403)"))
                                Console.Error.WriteLine("\nLooks like you are not a tenant administrator:\n\n");
                            Console.Error.WriteLine(e.ToString());
                            Console.Error.Write("\n\nUnable to get SharePoint context, press enter to close...");
                            Console.ReadLine();
                            Environment.Exit(0);
                        }
                    }
                    else
                    {
                        Environment.Exit(0);
                    }
                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine("\nError:\n\n" + e.ToString());
                Console.Error.Write("\n\nUnable to get SharePoint context, press enter to close...");
                Console.ReadLine();
                Environment.Exit(0);
            }
        }

        internal void Run()
        {
            Console.WriteLine("Loading sites information...\n");
            int sitesCount = GetSitesCount();
            Console.WriteLine("Found {0} sites in the tenant.", sitesCount);

            Console.Write("\nEnter the user CLAIM to get the permissions: ");
            string userClaim = Console.ReadLine();
            Console.Write("\nEnter a search query to filter the sites list against which the user permissions will be validated\n(e.g. */teams/wildcard*, leave empty if you want to validate all the sites): ");
            string siteSearchKey = Console.ReadLine();

            string fullPath = null;
            FileInfo fileInfo = null;
            while (true)
            {
                Console.Write("\nWhere do you want to save the results spreadsheet? (e.g. C:\\temp\\results.xlsx): ");
                fullPath = Console.ReadLine();
                fileInfo = new FileInfo(fullPath);

                if (!fileInfo.Exists) break;

                Console.Error.WriteLine("Oops! The file '{0}' already exists...\n", fullPath);
            }

            try
            {
                var perms = GetUserPermissions(userClaim, siteSearchKey);

                using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
                {
                    var workSheet = excelPackage.Workbook.Worksheets.Add("SitesPermissions");
                    workSheet.Cells["A1"].Value = "SiteUrl";
                    workSheet.Cells["B1"].Value = string.Format("Permissions for user {0}", userClaim);
                    workSheet.Cells["A1:B1"].Style.Font.Bold = true;
                    workSheet.Cells["A2"].LoadFromCollection(perms, false);
                    excelPackage.Save();
                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine("\n\nError:\n\n{0}", e.ToString());
                Console.Error.Write("\n\nPress enter to close...");
                Console.ReadLine();
                Environment.Exit(0);
            }

            Console.WriteLine("\n\nResults saved to {0}.", fullPath);

            Console.Write("\n\nPress enter to close...");
            Console.ReadLine();
        }

        internal int GetSitesCount()
        {
            return _svc.GetSites(null).Count();
        }

        internal IEnumerable<UserPermissions> GetUserPermissions(string user, string sitesSearchKey = null, bool broadlySharedOnly = true)
        {
            IEnumerable<SiteCollection> sites = _svc.GetSites(sitesSearchKey);

            sites = sites.OrderBy(s => s.Url).Take(sites.Count());
            List<UserPermissions> permissions = new List<UserPermissions>();

            var empty = new string[0];
            Console.WriteLine("\nValidating {0} sites, please wait...", sites.Count());
            TableBuilder tb = new TableBuilder();
            tb.AddRow("Site Url", "Permissions");
            tb.AddRow("--------", "-----------");
            for (int i = 0; i < sites.Count(); i++)
            {
                var site = sites.ElementAt(i);
                var perm = new UserPermissions() { SiteUrl = site.Url, /*PermissionsArray = empty,*/ Permissions = string.Empty };
                IEnumerable<string> perms = null;
                try
                {
                    var svc = new SiteService(site.Url, _clientId, _clientSecret);
                    perms = svc.GetUserPermissions(user);
                    perm.Permissions = perms != null ? string.Join(", ", perms) : string.Empty;
                }
                catch (Exception e)
                {
                    perm.Permissions = e.ToString();
                }
                finally
                {
                    if (broadlySharedOnly && perms.Count() > 1 || !broadlySharedOnly)
                    {
                        permissions.Add(perm);
                        int maxChars = 60;
                        string prm = new StringReader(perm.Permissions).ReadLine();
                        prm = prm.Length <= maxChars ? prm : prm.Substring(0, maxChars) + "...";
                        tb.AddRow(perm.SiteUrl, prm);
                    }
                }
            }

            Console.WriteLine(tb.Output());

            return permissions;
        }
    }

}
