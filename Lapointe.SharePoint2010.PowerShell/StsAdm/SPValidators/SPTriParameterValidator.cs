using System;
using Microsoft.SharePoint;

namespace Lapointe.SharePoint.PowerShell.StsAdm.SPValidators
{
    internal class SPTriParameterValidator
    {
        /// <summary>
        /// Validates the specified param A name.
        /// </summary>
        /// <param name="paramAName">Name of the param A.</param>
        /// <param name="paramAValue">The param A value.</param>
        /// <param name="paramBName">Name of the param B.</param>
        /// <param name="paramBValue">The param B value.</param>
        /// <param name="paramCName">Name of the param C.</param>
        /// <param name="paramCValue">The param C value.</param>
        public static void Validate(string paramAName, string paramAValue, string paramBName, string paramBValue, string paramCName, string paramCValue)
        {
            if ((paramAValue == null) && (paramBValue == null) && (paramCValue == null))
            {
                throw new ArgumentException(string.Format("Specify either the {0}, the {1}, or the {2} parameters with this command.", paramAName, paramBName, paramCName));
            }
            if ((paramAValue != null) && (paramBValue != null) && paramCValue != null)
                throw new ArgumentException(string.Format("The {0}, {1}, and {2} parameters are not compatible.  Please specify one or the other.", paramAName, paramBName, paramCName));
            if ((paramAValue != null) && (paramBValue != null))
            {
                throw new ArgumentException(SPResource.GetString("IncompatibleParametersSpecified", new object[] { paramAName, paramBName }));
            }
            if ((paramAValue != null) && (paramCValue != null))
            {
                throw new ArgumentException(SPResource.GetString("IncompatibleParametersSpecified", new object[] { paramAName, paramCName }));
            }
            if ((paramBValue != null) && (paramCValue != null))
            {
                throw new ArgumentException(SPResource.GetString("IncompatibleParametersSpecified", new object[] { paramBName, paramCName }));
            }
        }
    }

 

}
