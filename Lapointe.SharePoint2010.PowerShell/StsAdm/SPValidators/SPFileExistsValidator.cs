using System.IO;

namespace Lapointe.SharePoint.PowerShell.StsAdm.SPValidators
{
    internal class SPFileExistsValidator : SPNonEmptyValidator
    {
        /// <summary>
        /// Validates that the specified file exists.
        /// </summary>
        /// <param name="filename">The filename.</param>
        /// <returns></returns>
        public override bool Validate(string filename)
        {
            if (!base.Validate(filename))
            {
                return false;
            }
            return File.Exists(filename);
        }
    }
}
