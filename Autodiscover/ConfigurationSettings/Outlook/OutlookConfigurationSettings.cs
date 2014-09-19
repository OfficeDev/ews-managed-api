// ---------------------------------------------------------------------------
// <copyright file="OutlookConfigurationSettings.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the OutlookConfigurationSettings class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Represents Outlook configuration settings.
    /// </summary>
    internal sealed class OutlookConfigurationSettings : ConfigurationSettingsBase
    {
        #region Static fields
        /// <summary>
        /// All user settings that are available from the Outlook provider.
        /// </summary>
        private static LazyMember<List<UserSettingName>> allOutlookProviderSettings = new LazyMember<List<UserSettingName>>(
            () =>
            {
                List<UserSettingName> results = new List<UserSettingName>();
                results.AddRange(OutlookUser.AvailableUserSettings);
                results.AddRange(OutlookProtocol.AvailableUserSettings);
                results.Add(UserSettingName.AlternateMailboxes);
                return results;
            });
        #endregion

        #region Private fields
        private OutlookUser user;
        private OutlookAccount account;
        #endregion

        /// <summary>
        /// Initializes a new instance of the <see cref="OutlookConfigurationSettings"/> class.
        /// </summary>
        public OutlookConfigurationSettings()
        {
            this.user = new OutlookUser();
            this.account = new OutlookAccount();
        }

        /// <summary>
        /// Determines whether user setting is available in the OutlookConfiguration or not.
        /// </summary>
        /// <param name="setting">The setting.</param>
        /// <returns>True if user setting is available, otherwise, false.
        /// </returns>
        internal static bool IsAvailableUserSetting(UserSettingName setting)
        {
            return allOutlookProviderSettings.Member.Contains(setting);
        }

        /// <summary>
        /// Gets the namespace that defines the settings.
        /// </summary>
        /// <returns>The namespace that defines the settings.</returns>
        internal override string GetNamespace()
        {
            return "http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a";
        }

        /// <summary>
        /// Makes this instance a redirection response.
        /// </summary>
        /// <param name="redirectUrl">The redirect URL.</param>
        internal override void MakeRedirectionResponse(Uri redirectUrl)
        {
            this.account = new OutlookAccount()
            {
                RedirectTarget = redirectUrl.ToString(),
                ResponseType = AutodiscoverResponseType.RedirectUrl
            };
        }

        /// <summary>
        /// Tries to read the current XML element.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True is the current element was read, false otherwise.</returns>
        internal override bool TryReadCurrentXmlElement(EwsXmlReader reader)
        {
            if (!base.TryReadCurrentXmlElement(reader))
            {
                switch (reader.LocalName)
                {
                    case XmlElementNames.User:
                        this.user.LoadFromXml(reader);
                        return true;
                    case XmlElementNames.Account:
                        this.account.LoadFromXml(reader);
                        return true;
                    default:
                        reader.SkipCurrentElement();
                        return false;
                }
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Convert OutlookConfigurationSettings to GetUserSettings response.
        /// </summary>
        /// <param name="smtpAddress">SMTP address requested.</param>
        /// <param name="requestedSettings">The requested settings.</param>
        /// <returns>GetUserSettingsResponse</returns>
        internal override GetUserSettingsResponse ConvertSettings(string smtpAddress, List<UserSettingName> requestedSettings)
        {
            GetUserSettingsResponse response = new GetUserSettingsResponse();
            response.SmtpAddress = smtpAddress;

            if (this.Error != null)
            {
                response.ErrorCode = AutodiscoverErrorCode.InternalServerError;
                response.ErrorMessage = this.Error.Message;
            }
            else 
            {
                switch (this.ResponseType)
                {
                    case AutodiscoverResponseType.Success:
                        response.ErrorCode = AutodiscoverErrorCode.NoError;
                        response.ErrorMessage = string.Empty;
                        this.user.ConvertToUserSettings(requestedSettings, response);
                        this.account.ConvertToUserSettings(requestedSettings, response);
                        this.ReportUnsupportedSettings(requestedSettings, response);
                        break;
                    case AutodiscoverResponseType.Error:
                        response.ErrorCode = AutodiscoverErrorCode.InternalServerError;
                        response.ErrorMessage = Strings.InvalidAutodiscoverServiceResponse;
                        break;
                    case AutodiscoverResponseType.RedirectAddress:
                        response.ErrorCode = AutodiscoverErrorCode.RedirectAddress;
                        response.ErrorMessage = string.Empty;
                        response.RedirectTarget = this.RedirectTarget;
                        break;
                    case AutodiscoverResponseType.RedirectUrl:
                        response.ErrorCode = AutodiscoverErrorCode.RedirectUrl;
                        response.ErrorMessage = string.Empty;
                        response.RedirectTarget = this.RedirectTarget;
                        break;
                    default:
                        EwsUtilities.Assert(
                            false,
                            "OutlookConfigurationSettings.ConvertSettings",
                            "An unexpected error has occured. This code path should never be reached.");
                        break;
                }
            }
            return response;
        }

        /// <summary>
        /// Reports any requested user settings that aren't supported by the Outlook provider.
        /// </summary>
        /// <param name="requestedSettings">The requested settings.</param>
        /// <param name="response">The response.</param>
        private void ReportUnsupportedSettings(List<UserSettingName> requestedSettings, GetUserSettingsResponse response)
        {
            // In English: find settings listed in requestedSettings that are not supported by the Legacy provider.
            IEnumerable<UserSettingName> invalidSettingQuery = from setting in requestedSettings
                                                               where !OutlookConfigurationSettings.IsAvailableUserSetting(setting)
                                                               select setting;

            // Add any unsupported settings to the UserSettingsError collection.
            foreach (UserSettingName invalidSetting in invalidSettingQuery)
            {
                UserSettingError settingError = new UserSettingError()
                {
                    ErrorCode = AutodiscoverErrorCode.InvalidSetting,
                    SettingName = invalidSetting.ToString(),
                    ErrorMessage = string.Format(Strings.AutodiscoverInvalidSettingForOutlookProvider, invalidSetting.ToString())
                };
                response.UserSettingErrors.Add(settingError);
            }
        }

        /// <summary>
        /// Gets the type of the response.
        /// </summary>
        /// <value>The type of the response.</value>
        internal override AutodiscoverResponseType ResponseType
        {
            get
            {
                if (this.account != null)
                {
                    return this.account.ResponseType;
                }
                else
                {
                    return AutodiscoverResponseType.Error;
                }
            }
        }

        /// <summary>
        /// Gets the redirect target.
        /// </summary>
        internal override string RedirectTarget
        {
            get
            {
                return this.account.RedirectTarget;
            }
        }
    }
}
