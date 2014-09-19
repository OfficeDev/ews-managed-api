// ---------------------------------------------------------------------------
// <copyright file="CalendarFolder.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the CalendarFolder class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a folder containing appointments.
    /// </summary>
    [ServiceObjectDefinition(XmlElementNames.CalendarFolder)]
    public class CalendarFolder : Folder
    {
        /// <summary>
        /// Binds to an existing calendar folder and loads the specified set of properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the calendar folder.</param>
        /// <param name="id">The Id of the calendar folder to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>A CalendarFolder instance representing the calendar folder corresponding to the specified Id.</returns>
        public static new CalendarFolder Bind(
            ExchangeService service,
            FolderId id,
            PropertySet propertySet)
        {
            return service.BindToFolder<CalendarFolder>(id, propertySet);
        }

        /// <summary>
        /// Binds to an existing calendar folder and loads its first class properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the calendar folder.</param>
        /// <param name="id">The Id of the calendar folder to bind to.</param>
        /// <returns>A CalendarFolder instance representing the calendar folder corresponding to the specified Id.</returns>
        public static new CalendarFolder Bind(ExchangeService service, FolderId id)
        {
            return CalendarFolder.Bind(
                service,
                id,
                PropertySet.FirstClassProperties);
        }

        /// <summary>
        /// Binds to an existing calendar folder and loads the specified set of properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the calendar folder.</param>
        /// <param name="name">The name of the calendar folder to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>A CalendarFolder instance representing the calendar folder with the specified name.</returns>
        public static new CalendarFolder Bind(
            ExchangeService service,
            WellKnownFolderName name,
            PropertySet propertySet)
        {
            return CalendarFolder.Bind(
                service,
                new FolderId(name),
                propertySet);
        }

        /// <summary>
        /// Binds to an existing calendar folder and loads its first class properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the calendar folder.</param>
        /// <param name="name">The name of the calendar folder to bind to.</param>
        /// <returns>A CalendarFolder instance representing the calendar folder with the specified name.</returns>
        public static new CalendarFolder Bind(ExchangeService service, WellKnownFolderName name)
        {
            return CalendarFolder.Bind(
                service,
                new FolderId(name),
                PropertySet.FirstClassProperties);
        }

        /// <summary>
        /// Initializes an unsaved local instance of <see cref="CalendarFolder"/>. To bind to an existing calendar folder, use CalendarFolder.Bind() instead.
        /// </summary>
        /// <param name="service">The ExchangeService object to which the calendar folder will be bound.</param>
        public CalendarFolder(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Obtains a list of appointments by searching the contents of this folder and performing recurrence expansion
        /// for recurring appointments. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="view">The view controlling the range of appointments returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public FindItemsResults<Appointment> FindAppointments(CalendarView view)
        {
            EwsUtilities.ValidateParam(view, "view");

            ServiceResponseCollection<FindItemResponse<Appointment>> responses = this.InternalFindItems<Appointment>(
                (SearchFilter)null,
                view,
                null /* groupBy */);

            return responses[0].Results;
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
