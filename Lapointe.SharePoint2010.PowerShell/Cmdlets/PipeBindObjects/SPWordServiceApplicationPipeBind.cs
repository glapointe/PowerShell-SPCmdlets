using Microsoft.Office.Word.Server.Service;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.PowerShell;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects
{
    public sealed class SPWordServiceApplicationPipeBind : SPCmdletPipeBind<WordServiceApplication>
    {
        // Fields
        private Guid m_Id;
        private string m_Name;

        // Methods
        public SPWordServiceApplicationPipeBind(WordServiceApplication app)
        {
            this.m_Id = app.Id;
        }

        public SPWordServiceApplicationPipeBind(Guid guid)
        {
            this.m_Id = guid;
        }

        public SPWordServiceApplicationPipeBind(string name)
        {
            if (!string.IsNullOrEmpty(name))
            {
                try
                {
                    this.m_Id = new Guid(name);
                }
                catch (FormatException)
                {
                }
                catch (OverflowException)
                {
                }
                if (this.m_Id == Guid.Empty)
                {
                    this.m_Name = name;
                }
            }
        }

        protected override void Discover(WordServiceApplication instance)
        {
        }

        public override WordServiceApplication Read()
        {
            WordServiceApplication app = null;
            if (m_Id != Guid.Empty)
                app = Utilities.GetServiceApplication("Word Automation Services", m_Id) as WordServiceApplication;
            else if (!string.IsNullOrEmpty(m_Name))
                app = Utilities.GetServiceApplication("Word Automation Services", m_Name) as WordServiceApplication;

            
            if (app == null)
            {
                throw new InvalidOperationException("Unable to locate the Word Automation Service Application.");
            }
            return app;
        }

    
    }

 

}
