
namespace Lapointe.SharePoint.PowerShell.StsAdm.SPValidators
{
    internal class SPTrueFalseValidator : SPValidator
    {
        /// <summary>
        /// Validates the specified string.
        /// </summary>
        /// <param name="str">The string to validate.</param>
        /// <returns></returns>
        public override bool Validate(string str)
        {
            if (!string.IsNullOrEmpty(str))
            {
                if ((str.CompareTo("true") != 0) && (str.CompareTo("false") != 0))
                {
                    return false;
                }
                return true;
            }
            return false;
        }
    }


}
