using System.Collections;
using System.IO;
using System.Xml;
using Microsoft.SharePoint.PowerShell;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects
{
    public sealed class PropertiesPipeBind : SPCmdletPipeBind<Hashtable>
    {
        private string _xml;
        private Hashtable _hash;

        public PropertiesPipeBind(Hashtable instance)
            : base(instance)
        {
            _hash = instance.Clone() as Hashtable;
        }
        public PropertiesPipeBind(XmlDocument instance)
        {
            _xml = instance.OuterXml;
        }

        public PropertiesPipeBind(string inputString)
        {
            XmlDocument xml = new XmlDocument();
            try
            {
                if (File.Exists(inputString))
                {
                    xml.Load(inputString);
                }
                else
                {
                    xml.LoadXml(inputString);
                }
            }
            catch
            {
                throw new SPCmdletPipeBindException("The input string is not a valid XML file.");
            }
            _xml = xml.OuterXml;
        }



        protected override void Discover(Hashtable instance)
        {
            _hash = instance.Clone() as Hashtable;
        }

        public override Hashtable Read()
        {
            if (_hash != null)
                return _hash;

            Hashtable props = new Hashtable();
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(_xml);
            if (xmlDoc.DocumentElement == null)
                return props;

            foreach (XmlElement propElement in xmlDoc.DocumentElement.ChildNodes)
            {
                props.Add(propElement.Attributes["Name"].Value, propElement.InnerText.Trim());
            }
            return props;
        }

    }

}
