using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace O365.Documentor.Inventory.Configuration
{
    public class ConfigurationManager
    {
        private ConfigurationManager()
        {}

        public static ConfigurationManager _instance = null;
        public static ConfigurationManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new ConfigurationManager();
                }
                return _instance;
            }
        }

        System.Net.ICredentials _UserCredential = null;
        public System.Net.ICredentials UserCredential
        {
            get {
                var userId = System.Configuration.ConfigurationManager.AppSettings.Get("UserId");
                var password = System.Configuration.ConfigurationManager.AppSettings.Get("Password");
                SecureString sPassword = new SecureString();
                password.ToList().ForEach(sPassword.AppendChar);
                if (_UserCredential == null) {
                    _UserCredential = new SharePointOnlineCredentials(userId, sPassword);
                }
                return _UserCredential;
            }
        }

        public string SiteCollectionUrl
        {
            get
            {
                return System.Configuration.ConfigurationManager.AppSettings.Get("SiteCollectionURL");
            }
        }

        public string OutputDirectory
        {
            get
            {
                return System.Configuration.ConfigurationManager.AppSettings.Get("OutputDirectory");
            }
        }

        public string ContentTypeGroupName
        {
            get
            {
                return System.Configuration.ConfigurationManager.AppSettings.Get("ContentTypeGroupName");
            }
        }

        public string ColumnGroupName
        {
            get
            {
                return System.Configuration.ConfigurationManager.AppSettings.Get("ColumnGroupName");
            }
        }
    }
}
