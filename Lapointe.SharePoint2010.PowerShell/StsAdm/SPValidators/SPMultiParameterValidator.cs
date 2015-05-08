using System;
using System.Collections.Generic;
using Microsoft.SharePoint;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;

namespace Lapointe.SharePoint.PowerShell.StsAdm.SPValidators
{
    internal class SPMultiParameterValidator
    {
        /// <summary>
        /// Validates the specified param A name.
        /// </summary>
        /// <param name="parameters">The parameters.</param>
        /// <param name="possibleParams">The possible params.</param>
        /// <param name="minAllowed">The min allowed.</param>
        /// <param name="maxAllowed">The max allowed.</param>
        public static void Validate(SPParamCollection parameters, string[] possibleParams, int minAllowed, int maxAllowed)
        {
            List<SPParam> foundParams = new List<SPParam>();
            foreach (string param in possibleParams)
            {
                if (parameters[param.ToLowerInvariant()].UserTypedIn)
                    foundParams.Add(parameters[param.ToLowerInvariant()]);
            }
            if (foundParams.Count < minAllowed || foundParams.Count > maxAllowed)
            {
                if (minAllowed == maxAllowed)
                    throw new ArgumentException(
                        string.Format("You must specify only {0} of the following parameters: {1}", minAllowed,
                                      string.Join(", ", possibleParams)));
                else if (minAllowed != maxAllowed)
                {
                    throw new ArgumentException(
                        string.Format(
                            "You must specify at least {0} and no more than {1} of the following parameters: {2}",
                            minAllowed, maxAllowed, string.Join(", ", possibleParams)));
                }
            }
        }
    }

 

}
