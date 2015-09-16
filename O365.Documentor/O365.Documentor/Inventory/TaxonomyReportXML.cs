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
    public class TaxonomyReportXML : Inventory.InventoryBase
    {

        #region Singleton
        private TaxonomyReportXML() : base("Taxonomy.xml")
        {
            //Implement the initialization of your singleton
        }
        private readonly static TaxonomyReportXML _instance = new TaxonomyReportXML();

        public static TaxonomyReportXML Instance
        {
            get
            {
                return _instance;
            }
        }
        #endregion

        public override void Init()
        {
            WriteOutput("<?xml version=\"1.0\" encoding=\"utf-8\" ?>");
            WriteOutput("<Taxonomy>");
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
                WriteOutput("<Groups>");
                foreach (var group in store.Groups)
                {
                    if (group.Name.ToLower().Equals(ConfigurationManager.Instance.TaxonomyGroupNameFilter.ToLower()))
                    {
                        WriteOutput(string.Format("<Group Name=\"{0}\">", group.Name));
                        context.Load(group.TermSets, t => t.Include(ts => ts.Id, ts => ts.Name));
                        context.ExecuteQueryRetry();
                        WriteOutput("<Termsets>");
                        foreach (var termset in group.TermSets)
                        {
                            context.Load(termset.Terms,
                                ts => ts.Include(t => t.Id, t => t.Name));
                            context.ExecuteQueryRetry();
                            WriteOutput(string.Format("<Termset Name=\"{0}\" Id=\"{1}\">", termset.Name, termset.Id));
                            ProcessTerms(context, termset.Terms, 0);
                            WriteOutput("</Termset>");
                        }
                        WriteOutput("</Termsets>");

                        WriteOutput("</Group>");

                    }
                }
                WriteOutput("</Groups>");
                WriteOutput("</Taxonomy> ");
            }
        }

        private void ProcessTerms(ClientContext context, TermCollection terms, int indentLevel)
        {
            if (terms != null && terms.Count > 0)
            {
                WriteOutput("<Terms>");
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
                    WriteOutput(string.Format("<Term Name=\"{0}\" Id=\"{1}\"></Term>", term.Name, term.Id));
                    context.Load(term.Terms);
                    context.ExecuteQueryRetry();
                    if (term.Terms != null && term.Terms.Count > 0)
                    {
                        ProcessTerms(context, term.Terms, (indentLevel + 1));
                    }
                }
                WriteOutput("</Terms>");
            }
        }
    }
}
