using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using O365.Documentor.Inventory.Configuration;

namespace O365.Documentor.Inventory
{
    public class SiteCollectionInventory : Inventory.InventoryBase
    {
        private SiteCollectionInventory() : base("SitesList.csv")
        {}
        #region Singleton
        private static SiteCollectionInventory _instance = null;
        public static SiteCollectionInventory Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new SiteCollectionInventory();
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
       
        public override void Init()
        {
            WriteOutput(string.Format("{0},{1},{2},{3},{4},{5},{6},{7}",
                    "Title",
                    "Url",
                    "Id",
                    "Parent Web Id",
                    "WebTemplate",
                    "Date Created",
                    "Last Item Modified Date",
                    "Master Url"
                    ));
        }

        public override void Execute(string url)
        {
            using (var context = GetContext(url))
            {
                var web = context.Web;
                context.Load(web);
                context.Load(web.ParentWeb);
                context.Load(web.Webs);
                context.ExecuteQuery();

                if (web.WebTemplate.ToLower() == "app") { return;  }

                Guid parentwebId = Guid.Empty;
                try
                {
                    parentwebId = web.ParentWeb.Id;
                }
                catch { }
                WriteOutput(string.Format("\"{0}\",\"{1}\",\"{2}\",\"{3}\",\"{4}\",\"{5}\",\"{6}\",\"{7}\"",
                        web.Title,
                        web.Url,
                        web.Id,
                        parentwebId,
                        web.WebTemplate,
                        web.Created,
                        web.LastItemModifiedDate,
                        web.MasterUrl
                        ));

                foreach (var childWeb in web.Webs)
                {
                    context.Load(childWeb.ParentWeb);
                    context.ExecuteQuery();
                    Execute(childWeb.Url);
                }
            }
        }
    }
}
