
namespace Lapointe.SharePoint.PowerShell.StsAdm.SPValidators
{
    internal class SPLongRangeValidator : SPNonEmptyValidator
    {
        private long m_nLower;
        private long m_nUpper;

        public SPLongRangeValidator(long nLower, long nUpper)
        {
            m_nLower = nLower;
            m_nUpper = nUpper;
        }

        public override bool Validate(string strParam)
        {
            if (!base.Validate(strParam))
            {
                return false;
            }
            long num = long.Parse(strParam);
            if (m_nLower > num)
            {
                return false;
            }
            return (num <= m_nUpper);
        }
    }

 

}
