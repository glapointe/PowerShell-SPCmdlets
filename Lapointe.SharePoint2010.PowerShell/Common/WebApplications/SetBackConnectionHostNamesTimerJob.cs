using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;

namespace Lapointe.SharePoint.PowerShell.Common.WebApplications
{
    public class SetBackConnectionHostNamesTimerJob : SPJobDefinition
	{
		private const string JOB_NAME = "job-set-back-connection-host-names-";
        private const string KEY_USER = "userName";
        private const string KEY_PWD = "password";
        private const string KEY_URLS = "urls";

		private static readonly string jobId = Guid.NewGuid().ToString();

		public SetBackConnectionHostNamesTimerJob() : base() { }

		/// <summary>
		/// Initializes a new instance of the <see cref="SetBackConnectionHostNamesTimerJob"/> class.
		/// </summary>
        public SetBackConnectionHostNamesTimerJob(SPService service)
			: base(JOB_NAME + jobId, service, null, SPJobLockType.None)
		{
			Title = "Set BackConnectionHostNames Registry Key";
		}

		/// <summary>
		/// Executes the job definition.
		/// </summary>
		/// <param name="targetInstanceId">For target types of <see cref="T:Microsoft.SharePoint.Administration.SPContentDatabase"></see> this is the database ID of the content database being processed by the running job. This value is Guid.Empty for all other target types.</param>
		public override void Execute(Guid targetInstanceId)
		{
		    string user = Properties[KEY_USER] as string;
		    string password = Properties[KEY_PWD] as string;

            if (string.IsNullOrEmpty(user) || password == null)
                throw new ArgumentNullException("Username and password is required.");

            if (user.IndexOf('\\') < 0)
                throw new ArgumentException("Username must be in the form \"DOMAIN\\USER\"");

            IntPtr userHandle = new IntPtr(0);
            WindowsImpersonationContext impersonatedUser = null;
            try
            {

                bool returnValue = NativeMethods.LogonUser(
                  user.Split('\\')[1],
                  user.Split('\\')[0],
                  password,
                  NativeMethods.LOGON32_LOGON_INTERACTIVE,
                  NativeMethods.LOGON32_PROVIDER_DEFAULT,
                  ref userHandle
                  );

                if (!returnValue)
                {
                    throw new Exception("Invalid Username");
                }
                WindowsIdentity newId = new WindowsIdentity(userHandle);
                impersonatedUser = newId.Impersonate();

                List<string> urls = Properties[KEY_URLS] as List<string>;
                if (urls == null)
                    urls = Common.WebApplications.SetBackConnectionHostNames.GetUrls();
                Common.WebApplications.SetBackConnectionHostNames.SetBackConnectionRegKey(urls);

            }
            finally
            {
                if (impersonatedUser != null)
                    impersonatedUser.Undo();

                NativeMethods.CloseHandle(userHandle);
            }
		}

		/// <summary>
		/// Submits the job.
		/// </summary>
		public void SubmitJob(string user, string password, List<string> urls)
		{
		    Properties[KEY_USER] = user;
		    Properties[KEY_PWD] = password;
            Properties[KEY_URLS] = urls;
			Schedule = new SPOneTimeSchedule(DateTime.Now);
			Update();
		}
	}
}
