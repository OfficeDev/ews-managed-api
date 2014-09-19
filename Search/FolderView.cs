// ---------------------------------------------------------------------------
// <copyright file="FolderView.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the FolderView class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the view settings in a folder search operation.
    /// </summary>
    public sealed class FolderView : PagedView
    {
        private FolderTraversal traversal;

        /// <summary>
        /// Gets the name of the view XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetViewXmlElementName()
        {
            return XmlElementNames.IndexedPageFolderView;
        }

        /// <summary>
        /// Gets the name of the view json type.
        /// </summary>
        /// <returns></returns>
        internal override string GetViewJsonTypeName()
        {
            return "IndexedPageView";
        }

        /// <summary>
        /// Gets the type of service object this view applies to.
        /// </summary>
        /// <returns>A ServiceObjectType value.</returns>
        internal override ServiceObjectType GetServiceObjectType()
        {
            return ServiceObjectType.Folder;
        }

        /// <summary>
        /// Writes the attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteAttributeValue(XmlAttributeNames.Traversal, this.Traversal);
        }

        /// <summary>
        /// Adds the json properties.
        /// </summary>
        /// <param name="jsonRequest">The json request.</param>
        /// <param name="service">The service.</param>
        internal override void AddJsonProperties(JsonObject jsonRequest, ExchangeService service)
        {
            jsonRequest.Add(XmlAttributeNames.Traversal, this.Traversal);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FolderView"/> class.
        /// </summary>
        /// <param name="pageSize">The maximum number of elements the search operation should return.</param>
        public FolderView(int pageSize)
            : base(pageSize)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FolderView"/> class.
        /// </summary>
        /// <param name="pageSize">The maximum number of elements the search operation should return.</param>
        /// <param name="offset">The offset of the view from the base point.</param>
        public FolderView(int pageSize, int offset)
            : base(pageSize, offset)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FolderView"/> class.
        /// </summary>
        /// <param name="pageSize">The maximum number of elements the search operation should return.</param>
        /// <param name="offset">The offset of the view from the base point.</param>
        /// <param name="offsetBasePoint">The base point of the offset.</param>
        public FolderView(
            int pageSize,
            int offset,
            OffsetBasePoint offsetBasePoint)
            : base(pageSize, offset, offsetBasePoint)
        {
        }

        /// <summary>
        /// Gets or sets the search traversal mode. Defaults to FolderTraversal.Shallow.
        /// </summary>
        public FolderTraversal Traversal
        {
            get { return this.traversal; }
            set { this.traversal = value; }
        }
    }
}