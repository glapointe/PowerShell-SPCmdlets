using System.Collections;
using System.Collections.Specialized;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;

namespace Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers
{
    public class SPParam
    {
        private bool m_bEnabled;
        private bool m_bIsFlag;
        private bool m_bIsRequired;
        private bool m_bUserTypedIn;
        private string m_strDefaultValue;
        private string m_strHelpMessage;
        private string m_strName;
        private string m_strShortName;
        private string m_strValue;
        private ISPValidator m_Validator;
        private ArrayList m_Validators;

        /// <summary>
        /// Initializes a new instance of the <see cref="SPParam"/> class.
        /// </summary>
        /// <param name="strName">Name of the STR.</param>
        /// <param name="strShortName">Short name of the STR.</param>
        public SPParam(string strName, string strShortName)
        {
            m_strName = strName;
            m_strShortName = strShortName;
            m_bIsFlag = true;
            m_bEnabled = true;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SPParam"/> class.
        /// </summary>
        /// <param name="strName">Name of the STR.</param>
        /// <param name="strShortName">Short name of the STR.</param>
        /// <param name="bIsRequired">if set to <c>true</c> [b is required].</param>
        /// <param name="strDefaultValue">The STR default value.</param>
        /// <param name="validator">The validator.</param>
        public SPParam(string strName, string strShortName, bool bIsRequired, string strDefaultValue, ISPValidator validator) : this(strName, strShortName, bIsRequired, strDefaultValue, validator, "")
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SPParam"/> class.
        /// </summary>
        /// <param name="strName">Name of the STR.</param>
        /// <param name="strShortName">Short name of the STR.</param>
        /// <param name="bIsRequired">if set to <c>true</c> [b is required].</param>
        /// <param name="strDefaultValue">The STR default value.</param>
        /// <param name="validator">The validator.</param>
        /// <param name="strHelpMessage">The STR help message.</param>
        public SPParam(string strName, string strShortName, bool bIsRequired, string strDefaultValue, ISPValidator validator, string strHelpMessage)
        {
            m_strName = strName;
            m_strShortName = strShortName;
            m_bIsRequired = bIsRequired;
            m_strDefaultValue = strDefaultValue;
            m_Validator = validator;
            m_strHelpMessage = strHelpMessage;
            m_bEnabled = true;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SPParam"/> class.
        /// </summary>
        /// <param name="strName">Name of the STR.</param>
        /// <param name="strShortName">Short name of the STR.</param>
        /// <param name="bIsRequired">if set to <c>true</c> [b is required].</param>
        /// <param name="strDefaultValue">The STR default value.</param>
        /// <param name="validators">The validators.</param>
        /// <param name="strHelpMessage">The STR help message.</param>
        public SPParam(string strName, string strShortName, bool bIsRequired, string strDefaultValue, ArrayList validators, string strHelpMessage)
        {
            m_strName = strName;
            m_strShortName = strShortName;
            m_bIsRequired = bIsRequired;
            m_strDefaultValue = strDefaultValue;
            m_strHelpMessage = strHelpMessage;
            m_bEnabled = true;
            m_Validators = validators;
        }

        /// <summary>
        /// Inits the value from.
        /// </summary>
        /// <param name="keyValues">The key values.</param>
        public void InitValueFrom(StringDictionary keyValues)
        {
            m_strValue = keyValues[Name];
            if (m_strValue == null)
            {
                m_strValue = keyValues[ShortName];
            }
            m_bUserTypedIn = m_strValue != null;
            if (m_strValue == string.Empty)
                m_strValue = null;
        }

        private string m_errorInfo = null;
        public virtual string ErrorInfo
        {
            get { return m_errorInfo; }
            private set { m_errorInfo = value; }
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        /// <returns></returns>
        public bool Validate()
        {
            if (!m_bIsFlag)
            {
                if (m_Validator == null)
                {
                    if (m_Validators != null)
                    {
                        foreach (ISPValidator validator in m_Validators)
                        {
                            if (!validator.Validate(Value))
                            {
                                ErrorInfo = validator.ErrorInfo;
                                return false;
                            }
                        }
                        return true;
                    }
                    else
                    {
                        return true;
                    }
                }
                if (!m_Validator.Validate(Value))
                {
                    ErrorInfo = m_Validator.ErrorInfo;
                    return false;
                }
            }
            else
            {
                return ((m_strValue == null) || (m_strValue.Trim().Length == 0));
            }
            return true;
        }

        /// <summary>
        /// Gets the default value.
        /// </summary>
        /// <value>The default value.</value>
        public string DefaultValue
        {
            get
            {
                return m_strDefaultValue;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="SPParam"/> is enabled.
        /// </summary>
        /// <value><c>true</c> if enabled; otherwise, <c>false</c>.</value>
        public bool Enabled
        {
            get
            {
                return m_bEnabled;
            }
            set
            {
                m_bEnabled = value;
            }
        }

        /// <summary>
        /// Gets the help message.
        /// </summary>
        /// <value>The help message.</value>
        public string HelpMessage
        {
            get
            {
                return m_strHelpMessage;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is flag.
        /// </summary>
        /// <value><c>true</c> if this instance is flag; otherwise, <c>false</c>.</value>
        public bool IsFlag
        {
            get
            {
                return m_bIsFlag;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether this instance is required.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance is required; otherwise, <c>false</c>.
        /// </value>
        public bool IsRequired
        {
            get
            {
                return m_bIsRequired;
            }
            set
            {
                m_bIsRequired = value;
            }
        }

        /// <summary>
        /// Gets the name.
        /// </summary>
        /// <value>The name.</value>
        public string Name
        {
            get
            {
                return m_strName;
            }
        }

        /// <summary>
        /// Gets the short name.
        /// </summary>
        /// <value>The short name.</value>
        public string ShortName
        {
            get
            {
                return m_strShortName;
            }
        }

        /// <summary>
        /// Gets a value indicating whether [user typed in].
        /// </summary>
        /// <value><c>true</c> if [user typed in]; otherwise, <c>false</c>.</value>
        public bool UserTypedIn
        {
            get
            {
                return m_bUserTypedIn;
            }
        }

        /// <summary>
        /// Gets the value.
        /// </summary>
        /// <value>The value.</value>
        public string Value
        {
            get
            {
                if (UserTypedIn)
                {
                    return m_strValue;
                }
                return m_strDefaultValue;
            }
        }
    }

 
 
}
