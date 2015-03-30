/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a search folder. 
    /// </summary>
    [ServiceObjectDefinition(XmlElementNames.SearchFolder)]
    public class SearchFolder : Folder
    {
        /// <summary>
        /// Binds to an existing search folder and loads the specified set of properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the search folder.</param>
        /// <param name="id">The Id of the search folder to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>A SearchFolder instance representing the search folder corresponding to the specified Id.</returns>
        public static new SearchFolder Bind(
            ExchangeService service,
            FolderId id,
            PropertySet propertySet)
        {
            return service.BindToFolder<SearchFolder>(id, propertySet);
        }

        /// <summary>
        /// Binds to an existing search folder and loads its first class properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the search folder.</param>
        /// <param name="id">The Id of the search folder to bind to.</param>
        /// <returns>A SearchFolder instance representing the search folder corresponding to the specified Id.</returns>
        public static new SearchFolder Bind(ExchangeService service, FolderId id)
        {
            return SearchFolder.Bind(
                service,
                id,
                PropertySet.FirstClassProperties);
        }

        /// <summary>
        /// Binds to an existing search folder and loads the specified set of properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the search folder.</param>
        /// <param name="name">The name of the search folder to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>A SearchFolder instance representing the search folder with the specified name.</returns>
        public static new SearchFolder Bind(
            ExchangeService service,
            WellKnownFolderName name,
            PropertySet propertySet)
        {
            return SearchFolder.Bind(
                service,
                new FolderId(name),
                propertySet);
        }

        /// <summary>
        /// Binds to an existing search folder and loads its first class properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the search folder.</param>
        /// <param name="name">The name of the search folder to bind to.</param>
        /// <returns>A SearchFolder instance representing the search folder with the specified name.</returns>
        public static new SearchFolder Bind(ExchangeService service, WellKnownFolderName name)
        {
            return SearchFolder.Bind(
                service,
                new FolderId(name),
                PropertySet.FirstClassProperties);
        }

        /// <summary>
        /// Initializes an unsaved local instance of <see cref="SearchFolder"/>. To bind to an existing search folder, use SearchFolder.Bind() instead.
        /// </summary>
        /// <param name="service">The ExchangeService object to which the search folder will be bound.</param>
        public SearchFolder(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Internal method to return the schema associated with this type of object.
        /// </summary>
        /// <returns>The schema associated with this type of object.</returns>
        internal override ServiceObjectSchema GetSchema()
        {
            return SearchFolderSchema.Instance;
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();

            if (this.SearchParameters != null)
            {
                this.SearchParameters.Validate();
            }
        }

        /// <summary>
        /// Gets the minimum required server version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this service object type is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2007_SP1;
        }

        #region Properties

        /// <summary>
        /// Gets the search parameters associated with the search folder.
        /// </summary>
        public SearchFolderParameters SearchParameters
        {
            get { return (SearchFolderParameters)this.PropertyBag[SearchFolderSchema.SearchParameters]; }
        }

        #endregion
    }
}