using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.PowerShell;

namespace Lapointe.SharePoint.PowerShell.Cmdlets
{
    [SPCmdlet]
    public abstract class SPGetCmdletBaseCustom<TCmdletObject> : SPGetCmdletBase<TCmdletObject> where TCmdletObject : class
    {
        protected override void InternalBeginProcessing()
        {
            base.InternalBeginProcessing();

            Logger.ExecutingCmdlet = this;
#if DEBUG
            bool debug = (this.MyInvocation.Line.IndexOf("-Debug", StringComparison.InvariantCultureIgnoreCase) >= 0);
            if (debug)
                System.Diagnostics.Debugger.Launch();
#endif
        }

    }
}
