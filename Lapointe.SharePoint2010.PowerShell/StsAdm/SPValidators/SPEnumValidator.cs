using System;
using System.Collections.Generic;

namespace Lapointe.SharePoint.PowerShell.StsAdm.SPValidators
{
    internal class SPEnumValidator : SPRegexValidator
    {
        private Type enumTypeValue;
        private string[] additionalValues = null;
        private string[] removeValues = null;

        /// <summary>
        /// Initializes a new instance of the <see cref="SPEnumValidator"/> class.
        /// </summary>
        /// <param name="enumType">Type of the enum.</param>
        public SPEnumValidator(Type enumType) : base(GetEnumRegex(enumType, null, null), SPRegexValidationType.Legal)
        {
            enumTypeValue = enumType;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SPEnumValidator"/> class.
        /// </summary>
        /// <param name="enumType">Type of the enum.</param>
        /// <param name="additionalValues">The additional values.</param>
        public SPEnumValidator(Type enumType, string[] additionalValues) : base(GetEnumRegex(enumType, additionalValues, null), SPRegexValidationType.Legal)
        {
            enumTypeValue = enumType;
            this.additionalValues = additionalValues;
        }

        public SPEnumValidator(Type enumType, string[] additionalValues, string[] removeValues)
            : base(GetEnumRegex(enumType, additionalValues, removeValues), SPRegexValidationType.Legal)
        {
            enumTypeValue = enumType;
            this.additionalValues = additionalValues;
            this.removeValues = removeValues;
        }

        /// <summary>
        /// Gets the enum regex.
        /// </summary>
        /// <param name="enumType">Type of the enum.</param>
        /// <param name="additionalValues">The additional values.</param>
        /// <param name="removeValues">The remove values.</param>
        /// <returns></returns>
        private static string GetEnumRegex(Type enumType, string[] additionalValues, string[] removeValues)
        {
            if (!enumType.IsEnum)
                throw new Exception("Type is not an enum type.");
            string[] names = Enum.GetNames(enumType);
            if (removeValues != null)
            {
                List<string> temp = new List<string>();
                foreach (string name in names)
                {
                    bool found = false;
                    foreach (string toRemove in removeValues)
                    {
                        if (name.ToLower() == toRemove.ToLower())
                        {
                            found = true;
                            break;
                        }
                    }
                    if (!found)
                        temp.Add(name);
                }
                names = temp.ToArray();
            }
            string regex = string.Join("$|^", names);
            if (additionalValues != null && additionalValues.Length > 0)
                regex += "$|^" + string.Join("$|^", additionalValues);

            regex = "(?i:^" + regex + "$)";
            return regex;
        }

        /// <summary>
        /// Gets the display value.
        /// </summary>
        /// <value>The display value.</value>
        internal string DisplayValue
        {
            get
            {
                string[] names = Enum.GetNames(enumTypeValue);
                if (removeValues != null)
                {
                    List<string> temp = new List<string>();
                    foreach (string name in names)
                    {
                        bool found = false;
                        foreach (string toRemove in removeValues)
                        {
                            if (name.ToLower() == toRemove.ToLower())
                            {
                                found = true;
                                break;
                            }
                        }
                        if (!found)
                            temp.Add(name);
                    }
                    names = temp.ToArray();
                }

                string val = string.Join(" | ", names);
                if (additionalValues != null && additionalValues.Length > 0)
                    val += " | " + string.Join(" | ", additionalValues);
                return val.ToLowerInvariant();
            }
        }
    }
}
