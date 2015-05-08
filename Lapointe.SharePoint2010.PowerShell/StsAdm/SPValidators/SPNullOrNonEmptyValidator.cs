namespace Lapointe.SharePoint.PowerShell.StsAdm.SPValidators
{
    internal class SPNullOrNonEmptyValidator : SPValidator
    {
        /// <summary>
        /// Validates the specified string.
        /// </summary>
        /// <param name="str">The string.</param>
        /// <returns></returns>
        public override bool Validate(string str)
        {
            if ((str != null) && (str.Trim().Length <= 0))
            {
                return false;
            }
            return true;
        }
    }

 

}
