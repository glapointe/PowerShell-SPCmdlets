using System;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Taxonomy;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects
{
    public sealed class SPTaxonomyGroupPipeBind : SPCmdletPipeBind<Group>
    {
        private TaxonomySession _taxSession;
        private Guid _termStoreId;
        private Guid _groupId;

        public SPTaxonomyGroupPipeBind(Group group)
        {
            _taxSession = Utilities.GetTaxonomySessionFromTermStore(group.TermStore);
            _termStoreId = group.TermStore.Id;
            _groupId = group.Id;
        }

        protected override void Discover(Group instance)
        {
            _taxSession = Utilities.GetTaxonomySessionFromTermStore(instance.TermStore);
            _termStoreId = instance.TermStore.Id;
            _groupId = instance.Id;
        }


        public override Group Read()
        {
            return _taxSession.TermStores[_termStoreId].Groups[_groupId];
        }
    }
}
