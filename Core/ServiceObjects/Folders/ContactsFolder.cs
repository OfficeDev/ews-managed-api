// ---------------------------------------------------------------------------
// <copyright file="ContactsFolder.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ContactsFolder class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a folder containing contacts.
    /// </summary>
    [ServiceObjectDefinition(XmlElementNames.ContactsFolder)]
    public class ContactsFolder : Folder
    {
        /// <summary>
        /// Initializes an unsaved local instance of <see cref="ContactsFolder"/>. To bind to an existing contacts folder, use ContactsFolder.Bind() instead.
        /// </summary>
        /// <param name="service">The ExchangeService object to which the contacts folder will be bound.</param>
        public ContactsFolder(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Binds to an existing contacts folder and loads the specified set of properties.
        /// </summary>
        /// <param name="service">The service to use to bind to the contacts folder.</param>
        /// <param name="id">The Id of the contacts folder to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>A ContactsFolder instance representing the contacts folder corresponding to the specified Id.</returns>
        public static new ContactsFolder Bind(
            ExchangeService service,
            FolderId id,
            PropertySet propertySet)
        {
            return service.BindToFolder<ContactsFolder>(id, propertySet);
        }

        /// <summary>
        /// Binds to an existing contacts folder and loads its first class properties.
        /// </summary>
        /// <param name="service">The service to use to bind to the contacts folder.</param>
        /// <param name="id">The Id of the contacts folder to bind to.</param>
        /// <returns>A ContactsFolder instance representing the contacts folder corresponding to the specified Id.</returns>
        public static new ContactsFolder Bind(ExchangeService service, FolderId id)
        {
            return ContactsFolder.Bind(
                service,
                id,
                PropertySet.FirstClassProperties);
        }

        /// <summary>
        /// Binds to an existing contacts folder and loads the specified set of properties.
        /// </summary>
        /// <param name="service">The service to use to bind to the contacts folder.</param>
        /// <param name="name">The name of the contacts folder to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>A ContactsFolder instance representing the contacts folder with the specified name.</returns>
        public static new ContactsFolder Bind(
            ExchangeService service,
            WellKnownFolderName name,
            PropertySet propertySet)
        {
            return ContactsFolder.Bind(
                service,
                new FolderId(name),
                propertySet);
        }

        /// <summary>
        /// Binds to an existing contacts folder and loads its first class properties.
        /// </summary>
        /// <param name="service">The service to use to bind to the contacts folder.</param>
        /// <param name="name">The name of the contacts folder to bind to.</param>
        /// <returns>A ContactsFolder instance representing the contacts folder with the specified name.</returns>
        public static new ContactsFolder Bind(ExchangeService service, WellKnownFolderName name)
        {
            return ContactsFolder.Bind(
                service,
                new FolderId(name),
                PropertySet.FirstClassProperties);
        }

        /// <summary>
        /// Gets the minimum required server version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this service object type is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2007_SP1;
        }
    }
}
