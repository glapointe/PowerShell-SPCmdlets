using System;

namespace Lapointe.SharePoint.PowerShell.StsAdm.SPValidators
{
    internal class SPUrlValidator : SPNonEmptyValidator
    {
        // Methods
        public override bool Validate(string str)
        {
            if (base.Validate(str))
            {
                bool flag;
                if (str.IndexOf('\\') >= 0)
                {
                    return false;
                }
                try
                {
                    Uri uri = new Uri(str);
                    if (uri.Fragment == "")
                    {
                        if (uri.Query == "")
                        {
                            if (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps)
                            {
                                return true;
                            }
                            return false;
                        }
                    }
                    flag = false;
                }
                catch (UriFormatException)
                {
                    flag = false;
                }
                return flag;
            }
            return false;
        }
    }

}
