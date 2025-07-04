using System.Collections.Generic;
using System.DirectoryServices.ActiveDirectory;

namespace ADReplStatus
{
    public class ADREPLDC
    {
        public bool DiscoveryIssues;
        public string DomainName;
        public string IsGC;
        public string IsRODC;
        public string Name;
        public List<ReplicationNeighbor> ReplicationPartners = new List<ReplicationNeighbor>();
        public string Site;
    }
}