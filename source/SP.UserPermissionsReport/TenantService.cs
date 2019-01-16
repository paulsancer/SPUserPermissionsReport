using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SP.UserPermissionsReport
{
    public class TenantService : ITenantService
    {
        private readonly AuthenticationManager _authMgr;
        private readonly ClientContext _ctx;
        private readonly Tenant _tenant;
        private readonly string _tenantUrl;
        private readonly string _clientId;
        private readonly string _clientSecret;
        private readonly string _allSitesCacheKey = "AllSitesCacheKey";

        public TenantService(string tenantUrl)
        {
            try
            {
                _authMgr = new AuthenticationManager();
                _ctx = _authMgr.GetWebLoginClientContext(tenantUrl);
                _tenant = new Tenant(_ctx);

                Site site = _tenant.GetSiteByUrl(_tenantUrl);
                _ctx.Load(site, s => s.Owner.Email);
                _ctx.ExecuteQuery();

                _tenantUrl = tenantUrl;
            }
            catch (Exception e)
            {
                throw new Exception("TenantService.cs: " + e.ToString());
            }
        }

        public TenantService(string tenantUrl, string clientId, string clientSecret)
        {
            try
            {
                _authMgr = new AuthenticationManager();
                _ctx = _authMgr.GetAppOnlyAuthenticatedContext(tenantUrl, clientId, clientSecret);
                _tenant = new Tenant(_ctx);

                _tenantUrl = tenantUrl;
                _clientId = clientId;
                _clientSecret = clientSecret;
            }
            catch (Exception e)
            {
                throw new Exception("TenantService.cs: " + e.ToString());
            }
        }

        public IEnumerable<SiteCollection> GetAllSites()
        {
            SPOSitePropertiesEnumerable spp = null;
            int startIndex = 0;

            var sites = new List<SiteCollection>();

            while (spp == null || spp.Count > 0)
            {
                spp = _tenant.GetSiteProperties(startIndex, true);
                _ctx.Load(spp);
                _ctx.ExecuteQuery();


                foreach (SiteProperties sp in spp)
                    sites.Add(new SiteCollection()
                    {
                        Title = sp.Title,
                        Url = sp.Url,
                        Owner = sp.Owner
                    });

                startIndex += spp.Count;
            }

            return sites;
        }

        public IEnumerable<SiteCollection> GetSites(string searchKey)
        {
            IEnumerable<SiteCollection> sites = null;
            searchKey = searchKey ?? string.Empty;
            if (string.IsNullOrEmpty(searchKey) || searchKey.Contains("*"))
            {
                sites = MemoryCacher.Get(_allSitesCacheKey) as List<SiteCollection>;
                if (sites == null)
                {
                    sites = GetAllSites();
                    MemoryCacher.Add(_allSitesCacheKey, sites, DateTimeOffset.Now.AddHours(1));
                }
            }
            else
            {
                var site = GetSiteByUrl(searchKey);
                sites = site != null ? new SiteCollection[] { site } : new SiteCollection[0];
            }

            string key = searchKey.Replace("*", "");

            if (searchKey.StartsWith("*") && searchKey.EndsWith("*"))
                sites = sites.Where(s => s.Url.Contains(key));
            else if (searchKey.StartsWith("*"))
                sites = sites.Where(s => s.Url.EndsWith(key));
            else if (searchKey.EndsWith("*"))
                sites = sites.Where(s => s.Url.StartsWith(key));

            return sites;
        }

        //public IEnumerable<SiteCollection> GetAllSitesDetails(IEnumerable<SiteCollection> basicSites = null)
        //{
        //    basicSites = basicSites ?? GetAllSites();
        //    List<SiteCollectionDetails> sites = new List<SiteCollectionDetails>();
        //    foreach (var bSite in basicSites)
        //    {
        //        string admins = GetSiteAdministrators(bSite.Url);
        //        SiteCollectionDetails s = new SiteCollectionDetails() {
        //            Title = bSite.Title,
        //            Url = bSite.Url,
        //            Owner = bSite.Owner,
        //            Administrators = admins
        //        };
        //        sites.Add(s);
        //    }

        //    return sites;
        //}

        public string GetSiteAdministrators(string siteUrl)
        {
            try
            {
                AuthenticationManager authMgr = new AuthenticationManager();
                ClientContext ctx = authMgr.GetAppOnlyAuthenticatedContext(siteUrl, _clientId, _clientSecret);
                // Gets Site Collection Admins  
                List<UserEntity> admins = ctx.Site.RootWeb.GetAdministrators();
                StringBuilder sb = new StringBuilder();
                sb.Append("[");
                for (int i = 0; i < admins.Count; i++)
                {
                    sb.Append(string.IsNullOrEmpty(admins[i].Email) ? (string.IsNullOrEmpty(admins[i].LoginName) ? admins[i].LoginName : admins[i].Title) : admins[i].Email);
                    if (i < admins.Count + 1)
                        sb.Append("; ");
                }
                sb.Append("]");
                return sb.ToString();
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }


        public SiteCollection GetSiteByUrl(string siteUrl)
        {
            var tenant = new Tenant(_ctx);
            var siteCol = tenant.GetSiteByUrl(siteUrl);
            _ctx.Load(siteCol, s => s.Owner.Email);
            _ctx.ExecuteQuery();

            AuthenticationManager authMgr = new AuthenticationManager();
            ClientContext siteCtx = authMgr.GetAppOnlyAuthenticatedContext(siteUrl, _clientId, _clientSecret);
            var rootWeb = siteCtx.Web;
            siteCtx.Load(rootWeb, web => web.Title);
            siteCtx.ExecuteQuery();

            return new SiteCollection() { Title = rootWeb.Title, Url = siteUrl, Owner = siteCol.Owner.Email };
        }



        public bool IsUserEmailValid(string email)
        {
            try
            {
                User _newUser = _ctx.Web.EnsureUser(email);
                _ctx.Load(_newUser);

                _ctx.ExecuteQuery();

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
