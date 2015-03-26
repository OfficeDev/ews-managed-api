// ---------------------------------------------------------------------------
// <copyright file="UserSettingName.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    /// <summary>
    /// User settings that can be requested using GetUserSettings.
    /// </summary>
    /// <remarks>
    /// Add new values to the end and keep in sync with Microsoft.Exchange.Autodiscover.ConfigurationSettings.UserConfigurationSettingName.
    /// </remarks>
    public enum UserSettingName
    {
        /// <summary>
        /// The display name of the user.
        /// </summary>
        UserDisplayName = 0,

        /// <summary>
        /// The legacy distinguished name of the user.
        /// </summary>
        UserDN = 1,

        /// <summary>
        /// The deployment Id of the user.
        /// </summary>
        UserDeploymentId = 2,

        /// <summary>
        /// The fully qualified domain name of the mailbox server.
        /// </summary>
        InternalMailboxServer = 3,

        /// <summary>
        /// The fully qualified domain name of the RPC client server.
        /// </summary>
        InternalRpcClientServer = 4,

        /// <summary>
        /// The legacy distinguished name of the mailbox server.
        /// </summary>
        InternalMailboxServerDN = 5,

        /// <summary>
        /// The internal URL of the Exchange Control Panel.
        /// </summary>
        InternalEcpUrl = 6,

        /// <summary>
        /// The internal URL of the Exchange Control Panel for VoiceMail Customization.
        /// </summary>
        InternalEcpVoicemailUrl = 7,

        /// <summary>
        /// The internal URL of the Exchange Control Panel for Email Subscriptions.
        /// </summary>
        InternalEcpEmailSubscriptionsUrl = 8,

        /// <summary>
        /// The internal URL of the Exchange Control Panel for Text Messaging.
        /// </summary>
        InternalEcpTextMessagingUrl = 9,

        /// <summary>
        /// The internal URL of the Exchange Control Panel for Delivery Reports.
        /// </summary>
        InternalEcpDeliveryReportUrl = 10,

        /// <summary>
        /// The internal URL of the Exchange Control Panel for RetentionPolicy Tags.
        /// </summary>
        InternalEcpRetentionPolicyTagsUrl = 11,

        /// <summary>
        /// The internal URL of the Exchange Control Panel for Publishing.
        /// </summary>
        InternalEcpPublishingUrl = 12,

        /// <summary>
        /// The internal URL of the Exchange Control Panel for photos.
        /// </summary>
        InternalEcpPhotoUrl = 13,

        /// <summary>
        /// The internal URL of the Exchange Control Panel for People Connect subscriptions.
        /// </summary>
        InternalEcpConnectUrl = 14,

        /// <summary>
        /// The internal URL of the Exchange Control Panel for Team Mailbox.
        /// </summary>
        InternalEcpTeamMailboxUrl = 15,

        /// <summary>
        /// The internal URL of the Exchange Control Panel for creating Team Mailbox.
        /// </summary>
        InternalEcpTeamMailboxCreatingUrl = 16,

        /// <summary>
        /// The internal URL of the Exchange Control Panel for editing Team Mailbox.
        /// </summary>
        InternalEcpTeamMailboxEditingUrl = 17,

        /// <summary>
        /// The internal URL of the Exchange Control Panel for hiding Team Mailbox.
        /// </summary>
        InternalEcpTeamMailboxHidingUrl = 18,

        /// <summary>
        /// The internal URL of the Exchange Control Panel for the extension installation.
        /// </summary>
        InternalEcpExtensionInstallationUrl = 19,

        /// <summary>
        /// The internal URL of the Exchange Web Services.
        /// </summary>
        InternalEwsUrl = 20,

        /// <summary>
        /// The internal URL of the Exchange Management Web Services.
        /// </summary>
        InternalEmwsUrl = 21,

        /// <summary>
        /// The internal URL of the Offline Address Book.
        /// </summary>
        InternalOABUrl = 22,

        /// <summary>
        /// The internal URL of the Photos service.
        /// </summary>
        InternalPhotosUrl = 23,

        /// <summary>
        /// The internal URL of the Unified Messaging services.
        /// </summary>
        InternalUMUrl = 24,

        /// <summary>
        /// The internal URLs of the Exchange web client.
        /// </summary>
        InternalWebClientUrls = 25,

        /// <summary>
        /// The distinguished name of the mailbox database of the user's mailbox.
        /// </summary>
        MailboxDN = 26,

        /// <summary>
        /// The name of the Public Folders server.
        /// </summary>
        PublicFolderServer = 27,

        /// <summary>
        /// The name of the Active Directory server.
        /// </summary>
        ActiveDirectoryServer = 28,

        /// <summary>
        /// The name of the RPC over HTTP server.
        /// </summary>
        ExternalMailboxServer = 29,

        /// <summary>
        /// Indicates whether the RPC over HTTP server requires SSL.
        /// </summary>
        ExternalMailboxServerRequiresSSL = 30,

        /// <summary>
        /// The authentication methods supported by the RPC over HTTP server.
        /// </summary>
        ExternalMailboxServerAuthenticationMethods = 31,

        /// <summary>
        /// The URL fragment of the Exchange Control Panel for VoiceMail Customization.
        /// </summary>
        EcpVoicemailUrlFragment = 32,

        /// <summary>
        /// The URL fragment of the Exchange Control Panel for Email Subscriptions.
        /// </summary>
        EcpEmailSubscriptionsUrlFragment = 33,

        /// <summary>
        /// The URL fragment of the Exchange Control Panel for Text Messaging.
        /// </summary>
        EcpTextMessagingUrlFragment = 34,

        /// <summary>
        /// The URL fragment of the Exchange Control Panel for Delivery Reports.
        /// </summary>
        EcpDeliveryReportUrlFragment = 35,

        /// <summary>
        /// The URL fragment of the Exchange Control Panel for RetentionPolicy Tags.
        /// </summary>
        EcpRetentionPolicyTagsUrlFragment = 36,

        /// <summary>
        /// The URL fragment of the Exchange Control Panel for Publishing.
        /// </summary>
        EcpPublishingUrlFragment = 37,

        /// <summary>
        /// The URL fragment of the Exchange Control Panel for photos.
        /// </summary>
        EcpPhotoUrlFragment = 38,

        /// <summary>
        /// The URL fragment of the Exchange Control Panel for People Connect.
        /// </summary>
        EcpConnectUrlFragment = 39,

        /// <summary>
        /// The URL fragment of the Exchange Control Panel for Team Mailbox.
        /// </summary>
        EcpTeamMailboxUrlFragment = 40,

        /// <summary>
        /// The URL fragment of the Exchange Control Panel for creating Team Mailbox.
        /// </summary>
        EcpTeamMailboxCreatingUrlFragment = 41,

        /// <summary>
        /// The URL fragment of the Exchange Control Panel for editing Team Mailbox.
        /// </summary>
        EcpTeamMailboxEditingUrlFragment = 42,

        /// <summary>
        /// The URL fragment of the Exchange Control Panel for installing extension.
        /// </summary>
        EcpExtensionInstallationUrlFragment = 43,

        /// <summary>
        /// The external URL of the Exchange Control Panel.
        /// </summary>
        ExternalEcpUrl = 44,

        /// <summary>
        /// The external URL of the Exchange Control Panel for VoiceMail Customization.
        /// </summary>
        ExternalEcpVoicemailUrl = 45,

        /// <summary>
        /// The external URL of the Exchange Control Panel for Email Subscriptions.
        /// </summary>
        ExternalEcpEmailSubscriptionsUrl = 46,

        /// <summary>
        /// The external URL of the Exchange Control Panel for Text Messaging.
        /// </summary>
        ExternalEcpTextMessagingUrl = 47,

        /// <summary>
        /// The external URL of the Exchange Control Panel for Delivery Reports.
        /// </summary>
        ExternalEcpDeliveryReportUrl = 48,

        /// <summary>
        /// The external URL of the Exchange Control Panel for RetentionPolicy Tags.
        /// </summary>
        ExternalEcpRetentionPolicyTagsUrl = 49,

        /// <summary>
        /// The external URL of the Exchange Control Panel for Publishing.
        /// </summary>
        ExternalEcpPublishingUrl = 50,

        /// <summary>
        /// The external URL of the Exchange Control Panel for photos.
        /// </summary>
        ExternalEcpPhotoUrl = 51,

        /// <summary>
        /// The external URL of the Exchange Control Panel for People Connect subscriptions.
        /// </summary>
        ExternalEcpConnectUrl = 52,

        /// <summary>
        /// The external URL of the Exchange Control Panel for Team Mailbox.
        /// </summary>
        ExternalEcpTeamMailboxUrl = 53,

        /// <summary>
        /// The external URL of the Exchange Control Panel for creating Team Mailbox.
        /// </summary>
        ExternalEcpTeamMailboxCreatingUrl = 54,

        /// <summary>
        /// The external URL of the Exchange Control Panel for editing Team Mailbox.
        /// </summary>
        ExternalEcpTeamMailboxEditingUrl = 55,

        /// <summary>
        /// The external URL of the Exchange Control Panel for hiding Team Mailbox.
        /// </summary>
        ExternalEcpTeamMailboxHidingUrl = 56,

        /// <summary>
        /// The external URL of the Exchange Control Panel for the extension installation.
        /// </summary>
        ExternalEcpExtensionInstallationUrl = 57,

        /// <summary>
        /// The external URL of the Exchange Web Services.
        /// </summary>
        ExternalEwsUrl = 58,

        /// <summary>
        /// The external URL of the Exchange Management Web Services.
        /// </summary>
        ExternalEmwsUrl = 59,

        /// <summary>
        /// The external URL of the Offline Address Book.
        /// </summary>
        ExternalOABUrl = 60,

        /// <summary>
        /// The external URL of the Photos service.
        /// </summary>
        ExternalPhotosUrl = 61,

        /// <summary>
        /// The external URL of the Unified Messaging services.
        /// </summary>
        ExternalUMUrl = 62,

        /// <summary>
        /// The external URLs of the Exchange web client.
        /// </summary>
        ExternalWebClientUrls = 63,

        /// <summary>
        /// Indicates that cross-organization sharing is enabled.
        /// </summary>
        CrossOrganizationSharingEnabled = 64,

        /// <summary>
        /// Collection of alternate mailboxes.
        /// </summary>
        AlternateMailboxes = 65,

        /// <summary>
        /// The version of the Client Access Server serving the request (e.g. 14.XX.YYY.ZZZ)
        /// </summary>
        CasVersion = 66,

        /// <summary>
        /// Comma-separated list of schema versions supported by Exchange Web Services. The schema version values
        /// will be the same as the values of the ExchangeServerVersion enumeration.
        /// </summary>
        EwsSupportedSchemas = 67,

        /// <summary>
        /// The internal connection settings list for pop protocol
        /// </summary>
        InternalPop3Connections = 68,

        /// <summary>
        /// The external connection settings list for pop protocol
        /// </summary>
        ExternalPop3Connections = 69,

        /// <summary>
        /// The internal connection settings list for imap4 protocol
        /// </summary>
        InternalImap4Connections = 70,

        /// <summary>
        /// The external connection settings list for imap4 protocol
        /// </summary>
        ExternalImap4Connections = 71,

        /// <summary>
        /// The internal connection settings list for smtp protocol
        /// </summary>
        InternalSmtpConnections = 72,

        /// <summary>
        /// The external connection settings list for smtp protocol
        /// </summary>
        ExternalSmtpConnections = 73,

        /// <summary>
        /// If set to "Off" then clients should not connect via this protocol.
        /// The protocol contents are for informational purposes only.
        /// </summary>
        InternalServerExclusiveConnect = 74,

        /// <summary>
        /// The version of the Exchange Web Services server ExternalEwsUrl is pointing to.
        /// </summary>
        ExternalEwsVersion = 75,

        /// <summary>
        /// Mobile Mailbox policy settings.
        /// </summary>
        MobileMailboxPolicy = 76,

        /// <summary>
        /// Document sharing locations and their settings.
        /// </summary>
        DocumentSharingLocations = 77,

        /// <summary>
        /// Whether the user account is an MSOnline account.
        /// </summary>
        UserMSOnline = 78,

        /// <summary>
        /// The authentication methods supported by the RPC client server.
        /// </summary>
        InternalMailboxServerAuthenticationMethods = 79,

        /// <summary>
        /// Version of the server hosting the user's mailbox.
        /// </summary>
        MailboxVersion = 80,

        /// <summary>
        /// Sharepoint MySite Host URL.
        /// </summary>
        SPMySiteHostURL = 81,

        /// <summary>
        /// Site mailbox creation URL in SharePoint.
        /// It's used by Outlook to create site mailbox from SharePoint directly.
        /// </summary>
        SiteMailboxCreationURL = 82,

        /// <summary>
        /// The FQDN of the server used for internal RPC/HTTP connectivity.
        /// </summary>
        InternalRpcHttpServer = 83,

        /// <summary>
        /// Indicates whether SSL is required for internal RPC/HTTP connectivity.
        /// </summary>
        InternalRpcHttpConnectivityRequiresSsl = 84,

        /// <summary>
        /// The authentication method used for internal RPC/HTTP connectivity.
        /// </summary>
        InternalRpcHttpAuthenticationMethod = 85,

        /// <summary>
        /// If set to "On" then clients should only connect via this protocol.
        /// </summary>
        ExternalServerExclusiveConnect = 86,

        /// <summary>
        /// If set, then clients can call the server via XTC
        /// </summary>
        ExchangeRpcUrl = 87,

        /// <summary>
        /// If set to false then clients should not show the GAL by default, but show the contact list.
        /// </summary>
        ShowGalAsDefaultView = 88,

        /// <summary>
        /// AutoDiscover Primary SMTP Address for the user.
        /// </summary>
        AutoDiscoverSMTPAddress = 89,

        /// <summary>
        /// The 'interop' external URL of the Exchange Web Services.
        /// By interop it means a URL to E14 (or later) server that can serve mailboxes
        /// that are hosted in downlevel server (E2K3 and earlier).
        /// </summary>
        InteropExternalEwsUrl = 90,

        /// <summary>
        /// Version of server InteropExternalEwsUrl is pointing to.
        /// </summary>
        InteropExternalEwsVersion = 91,

        /// <summary>
        /// Public Folder (Hierarchy) information
        /// </summary>
        PublicFolderInformation = 92,

        /// <summary>
        /// The version appropriate URL of the AutoDiscover service that should answer this query.
        /// </summary>
        RedirectUrl = 93,

        /// <summary>
        /// The URL of the Exchange Web Services for Office365 partners.
        /// </summary>
        EwsPartnerUrl = 94,

        /// <summary>
        /// SSL certificate name
        /// </summary>
        CertPrincipalName = 95,

        /// <summary>
        /// The grouping hint for certain clients.
        /// </summary>
        GroupingInformation = 96,

        /// <summary>
        /// Internal OutlookService URL
        /// </summary>
        InternalOutlookServiceUrl = 98,

        /// <summary>
        /// External OutlookService URL
        /// </summary>
        ExternalOutlookServiceUrl = 99
    }
}
