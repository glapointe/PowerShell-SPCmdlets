using System;

namespace Lapointe.SharePoint.PowerShell.StsAdm.SPValidators
{
    internal class SPGuidValidator : SPNonEmptyValidator
    {
        /// <summary>
        /// Validates the specified GUID.
        /// </summary>
        /// <param name="guid">The GUID.</param>
        /// <returns></returns>
        public override bool Validate(string guid)
        {
            if (base.Validate(guid))
            {
                try
                {
                    new Guid(guid);
                    return true;
                }
                catch (UriFormatException)
                {
                    return false;
                }
            }
            return false;
        }
    }
}
