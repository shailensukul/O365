using O365.Documentor.Inventory.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace O365.Documentor.Inventory
{
    public abstract class InventoryBase
    {
        protected InventoryBase(string fileName)
        {
            FileName = 
            (string.Format("{0}\\{1}", 
            ConfigurationManager.Instance.OutputDirectory, fileName));
            File.Delete(FileName);
        }
        protected string FileName
        {
            get;
            set;
        }

        public abstract void Init();
        public abstract void Execute(string arg);

        protected Microsoft.SharePoint.Client.ClientContext GetContext(string url)
        {
            var context = new Microsoft.SharePoint.Client.ClientContext(url);
            context.Credentials = ConfigurationManager.Instance.UserCredential;
            return context;
        }

        protected void WriteOutput(string line)
        {
            //Console.WriteLine(line);
            using (var sw = System.IO.File.AppendText(FileName))
            {
                sw.WriteLine(line);
            }
        }
    }
}
