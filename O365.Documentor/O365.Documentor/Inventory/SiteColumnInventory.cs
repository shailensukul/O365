using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace O365.Documentor.Inventory
{
    public class SiteColumnInventory : InventoryBase
    {
        private SiteColumnInventory() : base("SiteColumns.csv")
        { }

        #region Singleton
        private static SiteColumnInventory _instance = null;
        public static SiteColumnInventory Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new SiteColumnInventory();
                }
                return _instance;
            }
        }
        #endregion

        public override void Init()
        {
            WriteOutput(
                string.Format("{0},{1},{2},{3},{4},{5},{6}",
                "Id",
                "Name",
                "Internal Name",
                "Group",
                "Type Display Name",
                "Type Description",
                "Description"
                )
                );
        }

        public override void Execute(string url)
        {
            using (var context = GetContext(url))
            {
                var web = context.Web;
                context.Load(web);
                context.Load(web.Fields);
                context.ExecuteQuery();

                foreach (var field in web.Fields)
                {
                    if (field.Group.ToLower().Equals(
                        Configuration.ConfigurationManager.Instance.ColumnGroupName.ToLower()))
                    {
                        //Console.WriteLine(contentType.Name);
                        WriteOutput(
                            string.Format("\"{0}\",\"{1}\",\"{2}\",\"{3}\",\"{4}\",\"{5}\",\"{6}\"",
                            field.Id,
                            field.StaticName,
                            field.InternalName,
                            field.Group,
                            field.TypeDisplayName,
                            field.TypeShortDescription,
                            field.Description)
                            );
                    }
                }
            }

        }
    }
}