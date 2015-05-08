using System;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Taxonomy;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects
{
    public sealed class SPTaxonomyTermPipeBind : SPCmdletPipeBind<Term>
    {
        private TaxonomySession _taxSession;
        private Guid _termStoreId;
        private Guid _termId;

        public SPTaxonomyTermPipeBind(Term term)
        {
            _taxSession = Utilities.GetTaxonomySessionFromTermStore(term.TermStore);
            _termStoreId = term.TermStore.Id;
            _termId = term.Id;
        }

        protected override void Discover(Term instance)
        {
            _taxSession = Utilities.GetTaxonomySessionFromTermStore(instance.TermStore);
            _termStoreId = instance.TermStore.Id;
            _termId = instance.Id;
        }

        public override Term Read()
        {
            return _taxSession.TermStores[_termStoreId].GetTerm(_termId);
        }
    }
}
