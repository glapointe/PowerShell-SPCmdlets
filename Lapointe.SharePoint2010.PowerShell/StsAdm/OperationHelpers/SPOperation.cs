using System;
using System.Collections.Specialized;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using Microsoft.SharePoint;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers
{
    public abstract class SPOperation : ISPStsadmCommand
    {
        private SPParamCollection m_params;
        private string m_helpMessage;
        protected const string NOT_VALID_FOR_FOUNDATION = "This command is not valid for SharePoint Foundation.";



        // 98d3057cd9024c27b2007643c1 is a special hard coded name for a list that Microsoft uses to store the mapping
        // of URLs from v2 to v3 (maps the bucket urls to the new urls).
        protected const string UPGRADE_AREA_URL_LIST = "98d3057cd9024c27b2007643c1";



        /// <summary>
        /// Inits the specified parameters.
        /// </summary>
        /// <param name="parameters">The parameters.</param>
        /// <param name="helpMessage">The help message.</param>
        protected void Init(SPParamCollection parameters, string helpMessage)
        {
            helpMessage = helpMessage.TrimEnd(new char[] {'\r', '\n'});
#if DEBUG
            helpMessage += "\r\n\t[-debug]";
            parameters.Add(new SPParam("debug", "debug"));
#endif

            m_params = parameters;
            
            helpMessage +=
                "\r\n\r\n\r\nCopyright 2010 Gary Lapointe\r\n  > For more information on this command and others:\r\n  > http://stsadm.blogspot.com/\r\n  > Use of this command is at your own risk.\r\n  > Gary Lapointe assumes no liability.";
            m_helpMessage = helpMessage;
        }

        /// <summary>
        /// Inits the parameters.
        /// </summary>
        /// <param name="keyValues">The key values.</param>
        public virtual void InitParameters(StringDictionary keyValues)
        {
            foreach (SPParam param in Params)
            {
                param.InitValueFrom(keyValues);
            }
            Validate(keyValues);
#if DEBUG
            if (Params["debug"].UserTypedIn)
                Debugger.Launch();
#endif
        }


        /// <summary>
        /// Validates the specified key values.
        /// </summary>
        /// <param name="keyValues">The key values.</param>
        public virtual void Validate(StringDictionary keyValues)
        {
            string strMessage = null;

            foreach (string current in keyValues.Keys)
            {
                if (current != "o" && Params[current] == null)
                {
                    strMessage += string.Format("Command line error. Invalid parameter: {0}.\r\n", current);
                }
            }
            if (strMessage != null)
                throw new SPSyntaxException(strMessage);

            foreach (SPParam param in Params)
            {
                if (param.Enabled)
                {
                    if (param.IsRequired && !param.UserTypedIn)
                    {
                        strMessage += SPResource.GetString("MissRequiredArg", new object[] { param.Name }) + "\r\n";
                    }
                }
            }
            if (strMessage != null)
                throw new SPSyntaxException(strMessage);

            foreach (SPParam param in Params)
            {
                if (param.Enabled)
                {
                    if (param.UserTypedIn && !param.Validate())
                    {
                        strMessage += SPResource.GetString("InvalidArg", new object[] { param.Name });
                        if (!string.IsNullOrEmpty(param.ErrorInfo))
                            strMessage += string.Format(" ({0})", param.ErrorInfo);

                        if (!string.IsNullOrEmpty(param.HelpMessage))
                        {
                            strMessage += "\r\n\t" + param.HelpMessage + "\r\n";
                        }
                    }
                }
            }
            if (strMessage != null)
                throw new SPSyntaxException(strMessage);
        }

        // Properties
        public virtual string DisplayNameId
        {
            get
            {
                return null;
            }
        }

       

        public string HelpMessage
        {
            get
            {
                return m_helpMessage;
            }
        }

        protected internal SPParamCollection Params
        {
            get
            {
                return m_params;
            }
        }

        /// <summary>
        /// Executes the specified command.
        /// </summary>
        /// <param name="command">The command.</param>
        /// <param name="keyValues">The key values.</param>
        /// <param name="output">The output.</param>
        /// <returns></returns>
        public abstract int Execute(string command, StringDictionary keyValues, out string output);

        #region ISPStsadmCommand Members

        /// <summary>
        /// Gets the help message.
        /// </summary>
        /// <param name="command">The command.</param>
        /// <returns></returns>
        public abstract string GetHelpMessage(string command);

        /// <summary>
        /// Runs the specified command.
        /// </summary>
        /// <param name="command">The command.</param>
        /// <param name="keyValues">The key values.</param>
        /// <param name="output">The output.</param>
        /// <returns></returns>
        public virtual int Run(string command, StringDictionary keyValues, out string output)
        {
            output = string.Empty;
            try
            {
                InitParameters(keyValues);

                return Execute(command, keyValues, out output);
            }
            catch (TargetInvocationException ex)
            {
                if (ex.InnerException != null)
                    throw ex.InnerException;
                else
                    throw;
            }
            catch (SPSyntaxException ex)
            {
                output += ex.Message;
                return (int) ErrorCodes.SyntaxError;
            }
        }

        #endregion
    }

    internal class SPSyntaxException : ApplicationException
    {
        // Methods
        public SPSyntaxException(string strMessage)
            : base(strMessage)
        {
        }
    }

 

 
}
