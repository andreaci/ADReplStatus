using CredentialManagement;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ADReplStatus
{
    internal class ADHelper
    {
        public string gForestName { get; set; } = string.Empty;
        public List<ADREPLDC> gDCs { get; set; } = new List<ADREPLDC>();
        public readonly ConcurrentBag<ADREPLDC> discoveredDCs  = new ConcurrentBag<ADREPLDC>();

        internal void Init()
        {
            if (string.IsNullOrEmpty(gForestName))
            {
                try
                {
                    using (Forest forest = Forest.GetCurrentForest())
                    {
                        gForestName = forest.Name;
                    }
                }
                catch (Exception ex)
                {
                    String err = $"Unable to detect AD forest. You will need to manually enter the AD forest you wish to scan using the 'Manually Set Forest' button.\nThis happens on non-domain joined computers as well as hybrid or Azure AD domain-joined machines: {ex.Message}";
                    LoggingManager.Error(err);
                    throw new Exception(err);
                }
            }
        }

        internal void DoWork(BackgroundWorker backgroundWorker1) 
        {
            Forest forest = null;

            try
            {
                DirectoryContext forestContext;
                var credential = new Credential { Target = "ADCredentials" };
                credential.Load();

                if (credential != null)
                {
                    forestContext = new DirectoryContext(DirectoryContextType.Forest, gForestName, credential.Username, credential.Password);
                }
                else
                {
                    forestContext = new DirectoryContext(DirectoryContextType.Forest, gForestName);
                }

                forest = Forest.GetForest(forestContext);
            }
            catch (Exception ex)
            {
                backgroundWorker1.ReportProgress(0, $"ERROR: Unable to find AD forest: {gForestName} \n{ex.Message}\n");
                return;
            }

            DomainCollection domainCollection = forest.Domains;
            backgroundWorker1.ReportProgress(0, $"Found {domainCollection.Count} domains in forest {forest.Name}.");

            foreach (Domain domain in domainCollection)
            {
                Parallel.ForEach(domain.DomainControllers.Cast<DomainController>(), dc =>
                {
                    ADREPLDC adrepldc = new ADREPLDC { Name = dc.Name, DomainName = domain.Name };
                    bool discoveryIssues = false;

                    try
                    {
                        using (var cts = new CancellationTokenSource(TimeSpan.FromSeconds(30)))
                        {
                            adrepldc.Site = dc.SiteName;
                        }
                    }
                    catch (Exception ex)
                    {
                        backgroundWorker1.ReportProgress(0, $"Failed to contact DC {dc.Name} for site name: {ex.Message}");
                        adrepldc.Site = "Unknown";
                        discoveryIssues = true;
                    }

                    try
                    {
                        adrepldc.IsGC = dc.IsGlobalCatalog().ToString();
                    }
                    catch (Exception ex)
                    {
                        backgroundWorker1.ReportProgress(0, $"Failed to determine GC status for {dc.Name}: {ex.Message}");
                        adrepldc.IsGC = "Unknown";
                        discoveryIssues = true;
                    }

                    try
                    {
                        using (DirectoryEntry directoryEntry = new DirectoryEntry("LDAP://" + dc.Name))
                        using (DirectorySearcher search = new DirectorySearcher(directoryEntry))
                        {
                            search.ClientTimeout = TimeSpan.FromSeconds(30);
                            search.Filter = $"(samaccountname={dc.Name.Split('.')[0]}$)";
                            search.PropertiesToLoad.Add("msDS-isRODC");
                            SearchResult result = search.FindOne();
                            adrepldc.IsRODC = result?.Properties.Contains("msDS-isRODC") == true && (bool)result.Properties["msDS-isRODC"][0] ? "True" : "False";
                        }
                    }
                    catch (Exception ex)
                    {
                        backgroundWorker1.ReportProgress(0, $"Failed to determine RODC status for {dc.Name}: {ex.Message}");
                        adrepldc.IsRODC = "Unknown";
                        discoveryIssues = true;
                    }

                    if (!discoveryIssues)
                    {
                        try
                        {
                            foreach (ReplicationNeighbor partner in dc.GetAllReplicationNeighbors())
                            {
                                adrepldc.ReplicationPartners.Add(partner);

                                if(partner.LastSyncResult != 0)
                                {
                                    discoveryIssues = true;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            backgroundWorker1.ReportProgress(0, $"Failed to determine replication neighbors for {dc.Name}: {ex.Message}");
                            discoveryIssues = true;
                        }
                    }

                    adrepldc.DiscoveryIssues = discoveryIssues;
                    discoveredDCs.Add(adrepldc);
                    backgroundWorker1.ReportProgress(100, $"OnDiscoveredDCsUpdated");
                });
            }
        }
    }
}
 