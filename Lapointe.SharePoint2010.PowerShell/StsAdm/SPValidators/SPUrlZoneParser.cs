using System;
using System.Globalization;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;

namespace Lapointe.SharePoint.PowerShell.StsAdm.SPValidators
{
    internal class SPUrlZoneParser
    {
        /// <summary>
        /// Parses the specified URL zone.
        /// </summary>
        /// <param name="strUrlZone">The URL zone.</param>
        /// <returns></returns>
        public static SPUrlZone Parse(string strUrlZone)
        {
            strUrlZone = strUrlZone.Trim().ToLower(CultureInfo.InvariantCulture);
            string str = strUrlZone;
            if (str != null)
            {
                if (str == "default")
                {
                    return SPUrlZone.Default;
                }
                if (str == "intranet")
                {
                    return SPUrlZone.Intranet;
                }
                switch (str)
                {
                    case "internet":
                        return SPUrlZone.Internet;

                    case "extranet":
                        return SPUrlZone.Extranet;

                    default:
                        if (str == "custom")
                        {
                            return SPUrlZone.Custom;
                        }
                        break;
                }
            }
            throw new ArgumentException(SPResource.GetString("ZoneNotFound", new object[] { strUrlZone }));
        }
    }
}
