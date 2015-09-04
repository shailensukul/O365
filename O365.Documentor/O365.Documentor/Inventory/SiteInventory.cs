using Microsoft.SharePoint.Client;
using System;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using O365.Documentor.Inventory.Configuration;
using System.Web;

namespace O365.Documentor.Inventory
{
    public class SiteInventory : Inventory.InventoryBase
    {
        private SiteInventory() : base("Sites.xml")
        { }
        #region Singleton
        private static SiteInventory _instance = null;
        public static SiteInventory Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new SiteInventory();
                }
                return _instance;
            }
        }
        #endregion

        private System.Net.ICredentials Credential
        {
            get;
            set;
        }

        private string _FeatureNameFilter = string.Empty;
        private string FeatureNameFilter
        {
            get
            {
                if (_FeatureNameFilter == string.Empty)
                {
                    _FeatureNameFilter = System.Configuration.ConfigurationManager.AppSettings.Get("FeatureNameFilter").ToLower();
                }
                return _FeatureNameFilter;
            }
        }

        public override void Init()
        {
        }

        public override void Execute(string url)
        {
            using (var context = GetContext(url))
            {
                var web = context.Web;
                context.Load(web);
                context.Load(web.ParentWeb);
                context.Load(web.Features, fcol => fcol.Include(f => f.DisplayName, f => f.DefinitionId));
                context.Load(web.Webs);
                context.ExecuteQuery();
                Guid parentwebId = Guid.Empty;
                try
                {
                    parentwebId = web.ParentWeb.Id;
                }
                catch { }
                WriteOutput(string.Format("<SiteCollection Title=\"{0}\" Url=\"{1}\" Template=\"{2}\" Description=\"{3}\">",
                        System.Web.HttpUtility.HtmlEncode(web.Title),
                        new Uri(web.Url).Segments.Last(),
                        web.WebTemplate,
                        System.Web.HttpUtility.HtmlEncode(web.Description)
                        ));
                // Get a list of features
                if (web.Features.Count > 0)
                {
                    WriteOutput("<Features>");
                    foreach (var feature in web.Features)
                    {
                        if (feature.DisplayName.ToLower().Contains(FeatureNameFilter))
                        {
                            WriteOutput(
                            string.Format("<Feature Id=\"{0}\" Name=\"{1}\"></Feature>",
                            feature.DefinitionId,
                            feature.DisplayName)
                            );
                        }
                    }
                    WriteOutput("</Features>");
                }
                if (web.Webs.Count > 0)
                {
                    WriteOutput("<Sites>");
                    foreach (var childWeb in web.Webs)
                    {
                        context.Load(childWeb.ParentWeb);
                        context.ExecuteQuery();
                        Process(childWeb.Url);
                    }
                    WriteOutput("</Sites>");
                }
                WriteOutput("</SiteCollection>");
            }

        }

        public void Process(string url)
        {
            using (var context = GetContext(url))
            {
                var web = context.Web;
                context.Load(web);
                context.Load(web.Features, fcol => fcol.Include(f => f.DisplayName, f => f.DefinitionId));
                context.Load(web.ParentWeb);
                context.Load(web.Webs);
                context.ExecuteQuery();

                if (web.WebTemplate.ToLower() == "app") { return; }

                Guid parentwebId = Guid.Empty;
                try
                {
                    parentwebId = web.ParentWeb.Id;
                }
                catch { }
                WriteOutput(string.Format("<Site Title=\"{0}\" Url=\"{1}\" Template=\"{2}\" Description=\"{3}\">",
                        System.Web.HttpUtility.HtmlEncode(web.Title),
                        System.Web.HttpUtility.HtmlDecode(new Uri(web.Url).Segments.Last()),
                        web.WebTemplate,
                        System.Web.HttpUtility.HtmlEncode(web.Description)
                        ));

                // Create dummy groups
                WriteOutput("<Permissions>");

                WriteOutput("<Groups>");
                WriteOutput(string.Format("<Group Name=\"Visitors\">"));
                WriteOutput("<Users>");
                WriteOutput("<User Name=\"c:0!.s|forms%3amembership\"></User>");
                WriteOutput("</Users>");
                WriteOutput("</Group>");
                WriteOutput(string.Format("<Group Name=\"Members\">"));
                WriteOutput("<Users>");
                WriteOutput("</Users>");
                WriteOutput("</Group>");
                WriteOutput(string.Format("<Group Name=\"Owners\">"));
                WriteOutput("<Users>");
                WriteOutput("<User Name=\"OWITSystemsTeam@officeworks.com.au\"></User>");
                WriteOutput("</Users>");
                WriteOutput("</Group>");
                WriteOutput("</Groups>");

                WriteOutput("</Permissions>");

                // Get a list of features
                if (web.Features.Count > 0)
                {
                    WriteOutput("<Features>");
                    foreach (var feature in web.Features)
                    {
                        if (feature.DisplayName.ToLower().Contains(FeatureNameFilter))
                        {
                            WriteOutput(
                            string.Format("<Feature Id=\"{0}\" Name=\"{1}\"></Feature>",
                            feature.DefinitionId,
                            feature.DisplayName)
                            );
                        }
                    }
                    WriteOutput("</Features>");
                }
                if (web.Webs.Count > 0)
                {
                    WriteOutput("<Sites>");
                    foreach (var childWeb in web.Webs)
                    {
                        context.Load(childWeb.ParentWeb);
                        context.ExecuteQuery();
                        Process(childWeb.Url);
                    }
                    WriteOutput("</Sites>");
                }
                WriteOutput("</Site>");
            }

        }
    }
}
