using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace O365.Documentor.Inventory
{
    public class ContentTypeInventory : InventoryBase
    {
        private ContentTypeInventory() : base("ContentTypes.csv")
        { }
        #region Singleton
        private static ContentTypeInventory _instance = null;
        public static ContentTypeInventory Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new ContentTypeInventory();
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
                context.Load(web);
                context.Load(web.ContentTypes);
                context.ExecuteQuery();

                foreach (var contentType in web.ContentTypes)
                {
                    if (contentType.Group.ToLower().Equals(
                        Configuration.ConfigurationManager.Instance.ContentTypeGroupName.ToLower()))
                    {
                        var fields = contentType.Fields;
                        context.Load(fields);
                        context.ExecuteQuery();


                        //Console.WriteLine(contentType.Name);
                        WriteOutput(
                            string.Format("\"{0}\",\"{1}\",\"{2}\",\"{3}\"",
                            "Content Type",
                            contentType.Id,
                            contentType.Name,
                            contentType.Group,
                            contentType.Description)
                            );
                        foreach (var field in fields)
                        {
                            if (!field.Hidden)
                            {
                                WriteOutput(
                                    string.Format("\"{0}\",\"{1}\",\"{2}\",\"{3}\",\"{4}\",\"{5}\",\"{6}\"",
                                    "Field",
                                    field.Id,
                                    field.StaticName,
                                    field.Group,
                                    field.InternalName,
                                    field.TypeDisplayName,
                                    field.TypeShortDescription)
                                    );
                            }
                        }
                    }
                }
            }
        }
    }
}
