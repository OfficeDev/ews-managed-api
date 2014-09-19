// ---------------------------------------------------------------------------
// <copyright file="BasicAuthModuleForUTF8.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the BasicAuthModuleForUTF8 class.</summary>
//-----------------------------------------------------------------------
namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Text;

    /// <summary>
    /// Custom basic authentication module for non ascii user names
    /// </summary>
    public class BasicAuthModuleForUTF8 : IAuthenticationModule 
    {   
        private const string AuthenticationTypeName = "Basic";
        private static BasicAuthModuleForUTF8 authModule = null; 
        private static object lockObject = new object();  

        /// <summary>
        /// Instantiation
        /// </summary>
        public static void InstantiateIfNeeded()
        { 
            lock (lockObject) 
            { 
                if (authModule == null)  
                {  
                    authModule = new BasicAuthModuleForUTF8(); 
                }  
            }  
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="BasicAuthModuleForUTF8"/> class.
        /// </summary>
        private BasicAuthModuleForUTF8()  
        {  
            AuthenticationManager.Unregister(AuthenticationTypeName); 
            AuthenticationManager.Register(this);
        }  

        /// <summary>
        /// AuthenticationType property
        /// </summary>
        string IAuthenticationModule.AuthenticationType 
        { 
            get { return AuthenticationTypeName; } 
        }  

        /// <summary>
        /// CanPreAuthenticate property
        /// </summary>
        bool IAuthenticationModule.CanPreAuthenticate 
        { 
            get { return true; }  
        }  

        /// <summary>
        /// Use custom implementaion of basic auth if the authType is Basic
        /// </summary>
        /// <param name="challenge">challenge to verify if it is basic</param>
        /// <param name="request">web request</param>
        /// <param name="credentials">credential</param>
        /// <returns></returns>
        Authorization IAuthenticationModule.Authenticate(string challenge, WebRequest request, ICredentials credentials)
        {  
            HttpWebRequest httpWebRequest = request as HttpWebRequest;  
            if (httpWebRequest == null) 
            { 
                return null;  
            }  

            // Verify that the challenge is a Basic Challenge 
            if (challenge == null || !challenge.StartsWith(AuthenticationTypeName, StringComparison.OrdinalIgnoreCase))
            { 
                return null; 
            }  
            return this.Authenticate(httpWebRequest, credentials); 
        }  

        /// <summary>
        /// PreAuthenticate implementation
        /// </summary>
        /// <param name="request">web request</param>
        /// <param name="credentials">credential</param>
        /// <returns></returns>
        Authorization IAuthenticationModule.PreAuthenticate(WebRequest request, ICredentials credentials)
        { 
            HttpWebRequest httpWebRequest = request as HttpWebRequest; 
            if (httpWebRequest == null) 
            { 
                return null;  
            } 
            return this.Authenticate(httpWebRequest, credentials); 
        }
  
        /// <summary>
        /// Custom implementaion of basic auth for non ascii email address.
        /// This is very similar to the .Net's Basic/default Authenticate implementation in ...\Net\System\Net\_BasicClient.cs, the only differenece here is the UTF8 encoding part 
        /// </summary>
        /// <param name="httpWebRequest">httpweb request object</param>
        /// <param name="credentials">user credential</param>
        /// <returns></returns>
        private Authorization Authenticate(HttpWebRequest httpWebRequest, ICredentials credentials)
        { 
            if (credentials == null) 
            { 
                return null; 
            } 

            // Get the username and password from the credentials 
            NetworkCredential nc = credentials.GetCredential(httpWebRequest.RequestUri, AuthenticationTypeName); 
            if (nc == null) 
            {  
                return null; 
            }  

            ICredentialPolicy policy = AuthenticationManager.CredentialPolicy;  
            if (policy != null && !policy.ShouldSendCredential(httpWebRequest.RequestUri, httpWebRequest, nc, this)) 
            {  
                return null; 
            }

            string username = nc.UserName;
            string domain = nc.Domain;

            if (String.IsNullOrEmpty(username))
            {
                return null;
            }

            string basicTicket = (!String.IsNullOrEmpty(domain) ? (domain + "\\") : "") + username + ":" + nc.Password; 
            byte[] bytes = Encoding.UTF8.GetBytes(basicTicket);
            string responseHeader = AuthenticationTypeName + " " + Convert.ToBase64String(bytes);
            return new Authorization(responseHeader, true); 
        }
     } 
}
