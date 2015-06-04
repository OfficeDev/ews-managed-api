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
    using System.ComponentModel;
    using System.Text;

    /// <summary>
    /// Represents a view settings that support paging in a search operation.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public abstract class PagedView : ViewBase
    {
        private int pageSize;
        private OffsetBasePoint offsetBasePoint = OffsetBasePoint.Beginning;
        private int offset;

        /// <summary>
        /// Write to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void InternalWriteViewToXml(EwsServiceXmlWriter writer)
        {
            base.InternalWriteViewToXml(writer);

            writer.WriteAttributeValue(XmlAttributeNames.Offset, this.Offset);
            writer.WriteAttributeValue(XmlAttributeNames.BasePoint, this.OffsetBasePoint);
        }

        /// <summary>
        /// Gets the maximum number of items or folders the search operation should return.
        /// </summary>
        /// <returns>The maximum number of items or folders that should be returned by the search operation.</returns>
        internal override int? GetMaxEntriesReturned()
        {
            return this.PageSize;
        }

        /// <summary>
        /// Internals the write search settings to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="groupBy">The group by clause.</param>
        internal override void InternalWriteSearchSettingsToXml(EwsServiceXmlWriter writer, Grouping groupBy)
        {
            if (groupBy != null)
            {
                groupBy.WriteToXml(writer);
            }
        }

        /// <summary>
        /// Writes OrderBy property to XML.
        /// </summary>
        /// <param name="writer">The writer</param>
        internal override void WriteOrderByToXml(EwsServiceXmlWriter writer)
        {
            // No order by for paged view
        }

        /// <summary>
        /// Validates this view.
        /// </summary>
        /// <param name="request">The request using this view.</param>
        internal override void InternalValidate(ServiceRequestBase request)
        {
            base.InternalValidate(request);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PagedView"/> class.
        /// </summary>
        /// <param name="pageSize">The maximum number of elements the search operation should return.</param>
        internal PagedView(int pageSize)
            : base()
        {
            this.PageSize = pageSize;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PagedView"/> class.
        /// </summary>
        /// <param name="pageSize">The maximum number of elements the search operation should return.</param>
        /// <param name="offset">The offset of the view from the base point.</param>
        internal PagedView(int pageSize, int offset)
            : this(pageSize)
        {
            this.Offset = offset;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PagedView"/> class.
        /// </summary>
        /// <param name="pageSize">The maximum number of elements the search operation should return.</param>
        /// <param name="offset">The offset of the view from the base point.</param>
        /// <param name="offsetBasePoint">The base point of the offset.</param>
        internal PagedView(
            int pageSize,
            int offset,
            OffsetBasePoint offsetBasePoint)
            : this(pageSize, offset)
        {
            this.OffsetBasePoint = offsetBasePoint;
        }

        /// <summary>
        /// The maximum number of items or folders the search operation should return.
        /// </summary>
        public int PageSize
        {
            get
            {
                return this.pageSize;
            }

            set
            {
                if (value <= 0)
                {
                    throw new ArgumentException(Strings.ValueMustBeGreaterThanZero);
                }

                this.pageSize = value;
            }
        }

        /// <summary>
        /// Gets or sets the base point of the offset.
        /// </summary>
        public OffsetBasePoint OffsetBasePoint
        {
            get { return this.offsetBasePoint; }
            set { this.offsetBasePoint = value; }
        }

        /// <summary>
        /// Gets or sets the offset.
        /// </summary>
        public int Offset
        {
            get
            {
                return this.offset;
            }

            set
            {
                if (value >= 0)
                {
                    this.offset = value;
                }
                else
                {
                    throw new ArgumentException(Strings.OffsetMustBeGreaterThanZero);
                }
            }
        }
    }
}