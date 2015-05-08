using System.IO;
using Microsoft.SharePoint;

namespace Lapointe.SharePoint.PowerShell.StsAdm.SPValidators
{
    public interface ISPValidator
    {
        bool Validate(string str);
        string ErrorInfo { get; }
    }


    internal class SPValidator : ISPValidator
    {

        public virtual bool Validate(string str)
        {
            return true;
        }
        private string m_errorInfo = null;
        public virtual string ErrorInfo
        {
            get { return m_errorInfo; }
            protected set { m_errorInfo = value;  }
        }
    }


    internal class SPDirectoryExistsAndValidFileNameValidator : SPNonEmptyValidator
    {
        public override bool Validate(string str)
        {
            if (base.Validate(str))
            {
                try
                {
                    FileInfo info = new FileInfo(str);
                    if (info.Directory.Exists)
                    {
                        if (info.Name.EndsWith("\\") || info.Name.EndsWith("/"))
                        {
                            ErrorInfo = "Filename not specified";
                            throw new SPException(SPResource.GetString("StsadmReqFileName", new object[0]));
                        }

                        return true;
                    }
                    else
                    {
                        ErrorInfo = "Directory not found";
                        throw new DirectoryNotFoundException();
                    }
                }
                catch
                {
                    return false;
                }
            }
            return false;
        }
    }


}
