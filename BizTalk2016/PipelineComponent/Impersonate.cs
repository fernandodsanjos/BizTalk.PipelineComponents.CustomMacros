using System;
using System.Runtime.InteropServices;
using System.Security.Principal;

namespace BizTalkComponents.PipelineComponents
{
    /// <summary>
    /// Object to change the user authentication
    /// </summary>
    public class Impersonate : IDisposable
    {
        public enum LoginType
        {
            LOGON32_LOGON_INTERACTIVE = 2,
            LOGON32_LOGON_NETWORK = 3,
            LOGON32_LOGON_BATCH = 4,
            LOGON32_LOGON_SERVICE = 5,
            LOGON32_LOGON_UNLOCK = 7,
            LOGON32_LOGON_NETWORK_CLEARTEXT = 8,
            LOGON32_LOGON_NEW_CREDENTIALS = 9
        }

        public enum LoginProvider
        {
            LOGON32_PROVIDER_DEFAULT = 0,
            LOGON32_PROVIDER_WINNT35 = 1,
            LOGON32_PROVIDER_WINNT40 = 2,
            LOGON32_PROVIDER_WINNT50 = 3
        }
        /// <summary>
        /// Logon method (check athetification) from advapi32.dll
        /// </summary>
        /// <param name="lpszUserName"></param>
        /// <param name="lpszDomain"></param>
        /// <param name="lpszPassword"></param>
        /// <param name="dwLogonType"></param>
        /// <param name="dwLogonProvider"></param>
        /// <param name="phToken"></param>
        /// <returns></returns>
        [DllImport("advapi32.dll")]
        private static extern bool LogonUser(String lpszUserName,
        String lpszDomain,
        String lpszPassword,
        int dwLogonType,
        int dwLogonProvider,
        ref IntPtr phToken);

        /// <summary>
        /// Close
        /// </summary>
        /// <param name="handle"></param>
        /// <returns></returns>
        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        public static extern bool CloseHandle(IntPtr handle);

        private WindowsImpersonationContext _windowsImpersonationContext;
        private IntPtr _tokenHandle;

        /// <summary>
        /// Initialize a UserImpersonation
        /// </summary>
        /// <param name="userName">domain\user</param>
        /// <param name="passWord"></param>
        public Impersonate(string userName, string passWord)
        {
            string[] domainUser = userName.Split(new string[] { @"\" }, StringSplitOptions.None);
            string currentDomain = WindowsIdentity.GetCurrent().Name.Split(new string[] { @"\" }, StringSplitOptions.RemoveEmptyEntries)[0];

            LoginType loginType = Impersonate.LoginType.LOGON32_LOGON_NETWORK;
            LoginProvider loginProvider = Impersonate.LoginProvider.LOGON32_PROVIDER_DEFAULT;

            //If different domain
            if (String.Equals(domainUser[0], currentDomain,StringComparison.OrdinalIgnoreCase) == false)
            {
                 loginType = Impersonate.LoginType.LOGON32_LOGON_NEW_CREDENTIALS;
                 loginProvider = Impersonate.LoginProvider.LOGON32_PROVIDER_WINNT50;
            }

            bool returnValue = LogonUser(domainUser[1], domainUser[0], passWord,
                    (int)loginType, (int)loginProvider,
                    ref _tokenHandle);

            WindowsIdentity newId = new WindowsIdentity(_tokenHandle);
            _windowsImpersonationContext = newId.Impersonate();

        }
        
     



        #region IDisposable Members

        /// <summary>
        /// Dispose the UserImpersonation connection
        /// </summary>
        public void Dispose()
        {
            if (_windowsImpersonationContext != null)
                _windowsImpersonationContext.Undo();
            if (_tokenHandle != IntPtr.Zero)
                CloseHandle(_tokenHandle);
        }

        #endregion
    }
}