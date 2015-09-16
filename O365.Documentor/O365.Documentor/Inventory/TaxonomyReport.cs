using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using O365.Documentor.Inventory.Configuration;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.SharePoint.Client;

namespace O365.Documentor.Taxonomy
{
    public class TaxonomyReport : Inventory.InventoryBase
    {

        #region Singleton
        private TaxonomyReport() : base("Taxonomy.csv")
        {
            //Implement the initialization of your singleton
        }
        private readonly static TaxonomyReport _instance = new TaxonomyReport();

        public static TaxonomyReport Instance
        {
            get
            {
                return _instance;
            }
        }
        #endregion

        public override void Init()
        {
            WriteOutput(string.Format("{0},{1},{2},{3},{4},{5},{6},{7}",
                   "Group",
                   "Group Id",
                   "Termset",
                   "Termset Id",
                   "Term Level 0",
                   "Id",
                   "Term Level 1",
                   "Id"
                   ));
        }

        public override void Execute(string url)
        {
            using (var context = GetContext(url))
            {
                var web = context.Web;
                var tax = TaxonomySession.GetTaxonomySession(context);
                var store = tax.GetDefaultSiteCollectionTermStore();
                context.Load(store,
                    s => s.WorkingLanguage,
                    s => s.Id,
                    s => s.Groups.Include(
                        g => g.Id,
                        g => g.Name
                        ));
                context.ExecuteQueryRetry();

                foreach (var group in store.Groups)
                {
                    if (group.Name.ToLower().Equals(ConfigurationManager.Instance.TaxonomyGroupNameFilter.ToLower()))
                    {
                        context.Load(group.TermSets, t => t.Include(ts => ts.Id, ts => ts.Name));
                        context.ExecuteQueryRetry();
                        foreach (var termset in group.TermSets)
                        {
                            context.Load(termset.Terms,
                                ts => ts.Include(t => t.Id, t => t.Name));
                            context.ExecuteQueryRetry();
                            ProcessTerms(context, termset.Terms, 0);
                        }
                    }
                }
            }
        }

        private void ProcessTerms(ClientContext context, TermCollection terms, int indentLevel)
        {
            if (terms != null && terms.Count > 0)
            {
                foreach (var term in terms)
                {
                    context.Load(term, 
                        t => t.Id, 
                        t => t.Name,
                        t => t.TermSet.Id,
                        t => t.TermSet.Name,
                        t => t.TermSet.Group.Id,
                        t => t.TermSet.Group.Name
                        );
                    context.ExecuteQueryRetry();
                    WriteOutput(string.Format("{0},{1},{2},{3},{4}{5},{6}",
                                 term.TermSet.Group.Name,
                                 term.TermSet.Group.Id,
                                 term.TermSet.Name,
                                 term.TermSet.Id,
                                 string.Format("{0}", new String(',', (indentLevel * 2))),                                                         
                                 term.Name,
                                 term.Id
                                 ));
                    context.Load(term.Terms);
                    context.ExecuteQueryRetry();
                    if (term.Terms != null && term.Terms.Count > 0)
                    {
                        ProcessTerms(context, term.Terms, (indentLevel + 1));
                    }
                }
            }
        }
    }
}
