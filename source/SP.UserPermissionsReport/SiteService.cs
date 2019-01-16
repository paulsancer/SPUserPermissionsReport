using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System;
using System.Collections.Generic;

namespace SP.UserPermissionsReport
{
    public class SiteService : ISiteService
    {
        private readonly AuthenticationManager _authMgr;
        private readonly ClientContext _ctx;
        private readonly string _siteUrl;
        private readonly string _appId;
        private readonly string _appSecret;

        public SiteService(string siteUrl, string appId, string appSecret)
        {
            try
            {
                _authMgr = new AuthenticationManager();
                _ctx = _authMgr.GetAppOnlyAuthenticatedContext(siteUrl, appId, appSecret);
                _siteUrl = siteUrl;
                _appId = appId;
                _appSecret = appSecret;
            }
            catch (Exception e)
            {
                throw new Exception("SPService.cs: " + e.ToString());
            }
        }

        //public IEnumerable<CustomAction> GetCustomActions()
        //{
        //    Site site = _ctx.Site;
        //    _ctx.Load(site, x => x.ServerRelativeUrl);
        //    _ctx.ExecuteQuery();

        //    UserCustomActionCollection spActions = _ctx.Site.UserCustomActions;

        //    _ctx.Load(spActions);
        //    _ctx.ExecuteQuery();

        //    var actions = new List<CustomAction>();
        //    if (spActions.Count > 0)
        //        foreach (var ca in spActions)
        //        {
        //            var action = new CustomAction()
        //            {
        //                Id = ca.Id.ToString(),
        //                Name = ca.Name,
        //                Title = ca.Title,
        //                Description = ca.Description,
        //                Location = ca.Location,
        //                Group = ca.Group,
        //                Url = ca.Url,
        //                Sequence = ca.Sequence.ToString(),
        //                ScriptBlock = ca.ScriptBlock,
        //                ScriptSrc = ca.ScriptSrc,
        //                ImageUrl = ca.ImageUrl
        //            };
        //        }

        //    return actions;
        //}

        public bool IsUserValid(string user)
        {
            try
            {
                User _newUser = _ctx.Web.EnsureUser(user);
                _ctx.Load(_newUser);

                _ctx.ExecuteQuery();

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool DoesUserHaveAccess(string user)
        {
            Web rootWeb = _ctx.Site.RootWeb;
            _ctx.Load(rootWeb);
            _ctx.ExecuteQuery();

            BasePermissions bp = new BasePermissions();
            bp.Set(PermissionKind.EmptyMask);
            var perm = rootWeb.GetUserEffectivePermissions(user);
            _ctx.ExecuteQuery();

            bool hasAccess = perm.Value.Has(PermissionKind.Open);

            return hasAccess;
        }

        public IEnumerable<string> GetUserPermissions(string user)
        {
            List<string> permissions = new List<string>();
            try
            {
                Web rootWeb = _ctx.Site.RootWeb;
                _ctx.Load(rootWeb);
                _ctx.ExecuteQuery();

                var perm = rootWeb.GetUserEffectivePermissions(user);
                _ctx.ExecuteQuery();

                foreach (PermissionKind k in Enum.GetValues(typeof(PermissionKind)))
                {
                    if (perm.Value.Has(k))
                        permissions.Add(k.ToString());
                }

                return permissions;
            }
            catch (Exception)
            {
                return permissions;
            }
        }
    }
}
