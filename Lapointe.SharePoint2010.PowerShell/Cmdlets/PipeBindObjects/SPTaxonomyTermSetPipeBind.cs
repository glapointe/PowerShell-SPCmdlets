using System;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Taxonomy;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects
{
    public sealed class SPTaxonomyTermSetPipeBind : SPCmdletPipeBind<TermSet>
    {
        private TaxonomySession _taxSession;
        private Guid _termStoreId;
        private Guid _termSetId;

        public SPTaxonomyTermSetPipeBind(TermSet termSet)
        {
            _taxSession = Utilities.GetTaxonomySessionFromTermStore(termSet.TermStore);
            _termStoreId = termSet.TermStore.Id;
            _termSetId = termSet.Id;
        }

        protected override void Discover(TermSet instance)
        {
            _taxSession = Utilities.GetTaxonomySessionFromTermStore(instance.TermStore);
            _termStoreId = instance.TermStore.Id;
            _termSetId = instance.Id;
        }

        public override TermSet Read()
        {
            return _taxSession.TermStores[_termStoreId].GetTermSet(_termSetId);
        }
    }
}
