// ---------------------------------------------------------------------------
// <copyright file="TaskSuggestion.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the TaskSuggestion class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.IO;

    /// <summary>
    /// Represents an TaskSuggestion object.
    /// </summary>
    public sealed class TaskSuggestion : ExtractedEntity
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="TaskSuggestion"/> class.
        /// </summary>
        internal TaskSuggestion()
            : base()
        {
        }

        /// <summary>
        /// Gets the meeting suggestion TaskString.
        /// </summary>
        public string TaskString { get; internal set; }

        /// <summary>
        /// Gets the meeting suggestion Assignees.
        /// </summary>
        public EmailUserEntityCollection Assignees { get; internal set; }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.NlgTaskString:
                    this.TaskString = reader.ReadElementValue();
                    return true;

                case XmlElementNames.NlgAssignees:
                    this.Assignees = new EmailUserEntityCollection();
                    this.Assignees.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.NlgAssignees);
                    return true;
                
                default:
                    return base.TryReadElementFromXml(reader);
            }
        }
    }
}
