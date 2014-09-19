// ---------------------------------------------------------------------------
// <copyright file="TasksFolder.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the TasksFolder class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a folder containing task items.
    /// </summary>
    [ServiceObjectDefinition(XmlElementNames.TasksFolder)]
    public class TasksFolder : Folder
    {
        /// <summary>
        /// Initializes an unsaved local instance of <see cref="TasksFolder"/>. To bind to an existing tasks folder, use TasksFolder.Bind() instead.
        /// </summary>
        /// <param name="service">The ExchangeService object to which the tasks folder will be bound.</param>
        public TasksFolder(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Binds to an existing tasks folder and loads the specified set of properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the tasks folder.</param>
        /// <param name="id">The Id of the tasks folder to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>A TasksFolder instance representing the task folder corresponding to the specified Id.</returns>
        public static new TasksFolder Bind(
            ExchangeService service,
            FolderId id,
            PropertySet propertySet)
        {
            return service.BindToFolder<TasksFolder>(id, propertySet);
        }

        /// <summary>
        /// Binds to an existing tasks folder and loads its first class properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the tasks folder.</param>
        /// <param name="id">The Id of the tasks folder to bind to.</param>
        /// <returns>A TasksFolder instance representing the task folder corresponding to the specified Id.</returns>
        public static new TasksFolder Bind(ExchangeService service, FolderId id)
        {
            return TasksFolder.Bind(
                service,
                id,
                PropertySet.FirstClassProperties);
        }

        /// <summary>
        /// Binds to an existing tasks folder and loads the specified set of properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the tasks folder.</param>
        /// <param name="name">The name of the tasks folder to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>A TasksFolder instance representing the tasks folder with the specified name.</returns>
        public static new TasksFolder Bind(
            ExchangeService service,
            WellKnownFolderName name,
            PropertySet propertySet)
        {
            return TasksFolder.Bind(
                service,
                new FolderId(name),
                propertySet);
        }

        /// <summary>
        /// Binds to an existing tasks folder and loads its first class properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the tasks folder.</param>
        /// <param name="name">The name of the tasks folder to bind to.</param>
        /// <returns>A TasksFolder instance representing the tasks folder with the specified name.</returns>
        public static new TasksFolder Bind(ExchangeService service, WellKnownFolderName name)
        {
            return TasksFolder.Bind(
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
