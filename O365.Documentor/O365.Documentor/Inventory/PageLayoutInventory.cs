using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace O365.Documentor.Inventory
{
    public class PageLayoutInventory : InventoryBase
    {
        private PageLayoutInventory() : base("PageLayoutInventory.csv")
        { }
        #region Singleton
        private static PageLayoutInventory _instance = null;
        public static PageLayoutInventory Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new PageLayoutInventory();
                }
                return _instance;
            }
        }
        #endregion

        public override void Init()
        {
            WriteOutput(
                string.Format("{0},{1},{2},{3},{4}",
                "Type",
                "Id",
                "Name",
                "Group",
                "Description"
                )
                );
        }
        public override void Execute(string url)
        {
            using (var context = GetContext(url))
            {
                var web = context.Web;
                var props = web.AllProperties;
                //var mpg = web.Lists.GetByTitle("Master Page Gallery");

                var list = context.Site.GetCatalog((int)ListTemplateType.MasterPageCatalog);
                var fields = list.Fields;

                context.Load(web, w => w.AllProperties);
                context.Load(props); 
                context.Load(list);
                context.Load(fields);
                context.ExecuteQuery();

                var pl = props["__PageLayouts"];

                var fieldnames = new List<string>();
                foreach (var field in fields)
                {
                    fieldnames.Add(field.InternalName);
                }
                var query = CamlQuery.CreateAllItemsQuery(100, fieldnames.ToArray());
                var items = list.GetItems(query);
                
                context.Load(items);
                context.ExecuteQuery();

                foreach (var item in items)
                {
                    foreach (var field in fields)
                    {
                        if (
                            field.InternalName.ToLower().Equals("contenttype") ||
                            field.InternalName.ToLower().Equals("permmask") ||
                            field.InternalName.ToLower().Equals("linkcheckedouttitle")                            
                            )
                        { continue;  }
                        //Console.WriteLine(field.InternalName);
                        try
                        {
                            Console.WriteLine(string.Format("{0} : {1}", field.InternalName, item[field.InternalName]));
                        }
                        catch { }
                    }
                    var filename = item["FileLeafRef"];
                    if (filename.ToString().ToLower().Contains("officeworks"))
                    {
                        Console.WriteLine(item["Title"]);
                    }
                }
            }
        }
    }
}
