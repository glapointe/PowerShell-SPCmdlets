namespace Lapointe.SharePoint.PowerShell.StsAdm.SPValidators
{
    internal class SPZoneUrlValidator : SPValidator
    {
        /// <summary>
        /// Validates the specified string.
        /// </summary>
        /// <param name="str">The string.</param>
        /// <returns></returns>
        public override bool Validate(string str)
        {
            if (str != null)
            {
                while (str.Trim().Length != 0)
                {
                    try
                    {
                        SPUrlZoneParser.Parse(str);
                        return true;
                    }
                    catch
                    {
                        return false;
                    }
                }
            }
            return false;
        }
    }

 

}
