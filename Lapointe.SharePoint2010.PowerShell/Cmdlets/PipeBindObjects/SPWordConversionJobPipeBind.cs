using Microsoft.Office.Word.Server.Conversions;
using Microsoft.SharePoint.PowerShell;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects
{
    public sealed class SPWordConversionJobPipeBind : SPCmdletPipeBind<Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects.SPWordConversionJobPipeBind.JobId>
    {
        private Guid _id = Guid.Empty;

        public SPWordConversionJobPipeBind(Guid id)
        {
            _id = id;
        }
        public SPWordConversionJobPipeBind(ConversionJob job) : this(job.JobId) { }

        protected override void Discover(JobId instance)
        {
            _id = instance.ID;
        }

        public override JobId Read()
        {
            return new JobId(_id);
        }
        public class JobId
        {
            public JobId(Guid id) { ID = id; }
            public Guid ID { get; set; }
        }
    }
}
