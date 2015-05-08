using System;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Taxonomy;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects
{
    public sealed class SPTaxonomyTermStorePipeBind : SPCmdletPipeBind<TermStore>
    {
        private TaxonomySession _taxSession;
        private Guid _termStoreId;

        public SPTaxonomyTermStorePipeBind(TermStore termStore)
        {
            _taxSession = Utilities.GetTaxonomySessionFromTermStore(termStore);
            _termStoreId = termStore.Id;
        }

        protected override void Discover(TermStore instance)
        {
            _taxSession = Utilities.GetTaxonomySessionFromTermStore(instance);
            _termStoreId = instance.Id;
        }

        public override TermStore Read()
        {
            return _taxSession.TermStores[_termStoreId];
        }
    }
}
