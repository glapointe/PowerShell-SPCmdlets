using System;
using Microsoft.SharePoint;

namespace Lapointe.SharePoint.PowerShell.StsAdm.SPValidators
{
    internal class SPBinaryParameterValidator
    {
        /// <summary>
        /// Validates the specified param A name.
        /// </summary>
        /// <param name="paramAName">Name of the param A.</param>
        /// <param name="paramAValue">The param A value.</param>
        /// <param name="paramBName">Name of the param B.</param>
        /// <param name="paramBValue">The param B value.</param>
        public static void Validate(string paramAName, string paramAValue, string paramBName, string paramBValue)
        {
            if ((paramAValue == null) && (paramBValue == null))
            {
                throw new ArgumentException(SPResource.GetString("MissingBinaryParameter", new object[] { paramAName, paramBName }));
            }
            if ((paramAValue != null) && (paramBValue != null))
            {
                throw new ArgumentException(SPResource.GetString("IncompatibleParametersSpecified", new object[] { paramAName, paramBName }));
            }
        }
    }

 

}
