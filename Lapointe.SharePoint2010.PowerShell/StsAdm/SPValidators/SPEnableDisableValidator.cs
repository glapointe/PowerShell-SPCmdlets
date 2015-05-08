
namespace Lapointe.SharePoint.PowerShell.StsAdm.SPValidators
{
    internal class SPEnableDisableValidator : SPValidator
    {
        /// <summary>
        /// Validates the specified string.
        /// </summary>
        /// <param name="str">The string to validate.</param>
        /// <returns></returns>
        public override bool Validate(string str)
        {
            if ((str != null) && (str.Trim().Length != 0))
            {
                if ((str.CompareTo("enable") != 0) && (str.CompareTo("disable") != 0))
                {
                    return false;
                }
                return true;
            }
            return false;
        }
    }


}
