#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

//-----------------------------------------------------------------------
// <summary>Defines the Conflict class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a conflict in a meeting time suggestion.
    /// </summary>
    public sealed class Conflict : ComplexProperty
    {
        private ConflictType conflictType;
        private int numberOfMembers;
        private int numberOfMembersAvailable;
        private int numberOfMembersWithConflict;
        private int numberOfMembersWithNoData;
        private LegacyFreeBusyStatus freeBusyStatus;

        /// <summary>
        /// Initializes a new instance of the <see cref="Conflict"/> class.
        /// </summary>
        /// <param name="conflictType">The type of the conflict.</param>
        internal Conflict(ConflictType conflictType)
            : base()
        {
            this.conflictType = conflictType;
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if appropriate element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.NumberOfMembers:
                    this.numberOfMembers = reader.ReadElementValue<int>();
                    return true;
                case XmlElementNames.NumberOfMembersAvailable:
                    this.numberOfMembersAvailable = reader.ReadElementValue<int>();
                    return true;
                case XmlElementNames.NumberOfMembersWithConflict:
                    this.numberOfMembersWithConflict = reader.ReadElementValue<int>();
                    return true;
                case XmlElementNames.NumberOfMembersWithNoData:
                    this.numberOfMembersWithNoData = reader.ReadElementValue<int>();
                    return true;
                case XmlElementNames.BusyType:
                    this.freeBusyStatus = reader.ReadElementValue<LegacyFreeBusyStatus>();
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service"></param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            foreach (string key in jsonProperty.Keys)
            {
                switch (key)
                {
                    case XmlElementNames.NumberOfMembers:
                        this.numberOfMembers = jsonProperty.ReadAsInt(key);
                        break;
                    case XmlElementNames.NumberOfMembersAvailable:
                        this.numberOfMembersAvailable = jsonProperty.ReadAsInt(key);
                        break;
                    case XmlElementNames.NumberOfMembersWithConflict:
                        this.numberOfMembersWithConflict = jsonProperty.ReadAsInt(key);
                        break;
                    case XmlElementNames.NumberOfMembersWithNoData:
                        this.numberOfMembersWithNoData = jsonProperty.ReadAsInt(key);
                        break;
                    case XmlElementNames.BusyType:
                        this.freeBusyStatus = jsonProperty.ReadEnumValue<LegacyFreeBusyStatus>(key);
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Gets the type of the conflict.
        /// </summary>
        public ConflictType ConflictType
        {
            get { return this.conflictType; }
        }

        /// <summary>
        /// Gets the number of users, resources, and rooms in the conflicting group. The value of this property
        /// is only meaningful when ConflictType is equal to ConflictType.GroupConflict.
        /// </summary>
        public int NumberOfMembers
        {
            get { return this.numberOfMembers; }
        }

        /// <summary>
        /// Gets the number of members who are available (whose status is Free) in the conflicting group. The value
        /// of this property is only meaningful when ConflictType is equal to ConflictType.GroupConflict.
        /// </summary>
        public int NumberOfMembersAvailable
        {
            get { return this.numberOfMembersAvailable; }
        }

        /// <summary>
        /// Gets the number of members who have a conflict (whose status is Busy, OOF or Tentative) in the conflicting
        /// group. The value of this property is only meaningful when ConflictType is equal to ConflictType.GroupConflict.
        /// </summary>
        public int NumberOfMembersWithConflict
        {
            get { return this.numberOfMembersWithConflict; }
        }

        /// <summary>
        /// Gets the number of members who do not have published free/busy data in the conflicting group. The value
        /// of this property is only meaningful when ConflictType is equal to ConflictType.GroupConflict.
        /// </summary>
        public int NumberOfMembersWithNoData
        {
            get { return this.numberOfMembersWithNoData; }
        }

        /// <summary>
        /// Gets the free/busy status of the conflicting attendee. The value of this property is only meaningful when
        /// ConflictType is equal to ConflictType.IndividualAttendee.
        /// </summary>
        public LegacyFreeBusyStatus FreeBusyStatus
        {
            get { return this.freeBusyStatus; }
        }
    }
}
