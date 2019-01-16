using System.Collections.Generic;

namespace SP.UserPermissionsReport
{
    public interface ITenantService
    {
        SiteCollection GetSiteByUrl(string siteUrl);
        IEnumerable<SiteCollection> GetSites(string searchKey);
        IEnumerable<SiteCollection> GetAllSites();
        string GetSiteAdministrators(string siteUrl);
    }
}