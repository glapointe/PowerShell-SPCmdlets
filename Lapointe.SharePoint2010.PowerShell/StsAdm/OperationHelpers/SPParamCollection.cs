using System.Collections;

namespace Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers
{
    public class SPParamCollection : IEnumerable
    {
        // Fields
        private ArrayList m_Collection = new ArrayList();
        private Hashtable m_NameMap = new Hashtable();
        private Hashtable m_ShortNameMap = new Hashtable();

        // Methods
        public void Add(SPParam param)
        {
            m_Collection.Add(param);
            m_NameMap.Add(param.Name, param);
            m_ShortNameMap.Add(param.ShortName, param);
        }

        public IEnumerator GetEnumerator()
        {
            return m_Collection.GetEnumerator();
        }

        // Properties
        public int Count
        {
            get
            {
                return m_Collection.Count;
            }
        }

        public SPParam this[string strName]
        {
            get
            {
                SPParam param = (SPParam) m_NameMap[strName];
                if (param != null)
                {
                    return param;
                }
                return (SPParam) m_ShortNameMap[strName];
            }
        }

        public SPParam this[int index]
        {
            get
            {
                return (SPParam) m_Collection[index];
            }
        }
    }

 
 
}
