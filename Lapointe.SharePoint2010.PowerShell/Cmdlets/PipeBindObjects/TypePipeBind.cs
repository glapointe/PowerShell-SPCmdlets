using System;
using System.IO;
using System.Xml;
using Microsoft.SharePoint.PowerShell;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects
{
    public sealed class TypePipeBind : SPCmdletPipeBind<Type>
    {
        private Type _type;

        public TypePipeBind(Type instance)
            : base(instance)
        {
            _type = instance;
        }

        public TypePipeBind(string inputString)
        {
            try
            {
                _type = Type.GetType(inputString, true, true);
            }
            catch
            {
                throw new SPCmdletPipeBindException("The input string is an invalid or unknown type. Ensure the assembly associated with the type is loaded in memory.");
            }
        }



        protected override void Discover(Type instance)
        {
            _type = instance;
        }

        public override Type Read()
        {
            return _type;
        }
    }

}
