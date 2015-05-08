using System.Globalization;
using System.Text.RegularExpressions;

namespace Lapointe.SharePoint.PowerShell.StsAdm.SPValidators
{
    internal class SPRegexValidator : SPNonEmptyValidator
    {
        private string m_strRegex;
        private SPRegexValidationType m_validationType;

        /// <summary>
        /// Initializes a new instance of the <see cref="SPRegexValidator"/> class.
        /// </summary>
        /// <param name="strRegex">The regex.</param>
        public SPRegexValidator(string strRegex) : this(strRegex, SPRegexValidationType.Legal)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SPRegexValidator"/> class.
        /// </summary>
        /// <param name="strRegex">The regex.</param>
        /// <param name="validationType">Type of the validation.</param>
        public SPRegexValidator(string strRegex, SPRegexValidationType validationType)
        {
            m_strRegex = strRegex;
            m_validationType = validationType;
        }

        /// <summary>
        /// Validates the specified parameter value.
        /// </summary>
        /// <param name="parameterValue">The parameter value.</param>
        /// <returns></returns>
        public override bool Validate(string parameterValue)
        {
            if (!base.Validate(parameterValue))
            {
                return false;
            }
            Regex regex = new Regex(m_strRegex);
            parameterValue = parameterValue.ToLower(CultureInfo.InvariantCulture);
            bool success = regex.Match(parameterValue).Success;
            if (m_validationType != SPRegexValidationType.Legal)
            {
                if (success)
                {
                    return false;
                }
                return base.Validate(parameterValue);
            }
            if (!success)
            {
                return false;
            }
            return base.Validate(parameterValue);
        }
    }

    internal enum SPRegexValidationType
    {
        Legal,
        Illegal
    }
}
