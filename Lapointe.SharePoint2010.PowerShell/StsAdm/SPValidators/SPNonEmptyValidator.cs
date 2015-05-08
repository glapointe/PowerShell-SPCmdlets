namespace Lapointe.SharePoint.PowerShell.StsAdm.SPValidators
{
    internal class SPNonEmptyValidator : SPValidator
    {
        /// <summary>
        /// Validates the specified string.
        /// </summary>
        /// <param name="str">The string.</param>
        /// <returns></returns>
        public override bool Validate(string str)
        {
            return ((str != null) && (str.Trim().Length != 0));
        }
    }
}
