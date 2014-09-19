// ---------------------------------------------------------------------------
// <copyright file="AbsoluteDateTransition.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AbsoluteDateTransition class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Text;

    /// <summary>
    /// Represents a time zone period transition that occurs on a fixed (absolute) date.
    /// </summary>
    internal class AbsoluteDateTransition : TimeZoneTransition
    {
        private DateTime dateTime;

        /// <summary>
        /// Initializes this transition based on the specified transition time.
        /// </summary>
        /// <param name="transitionTime">The transition time to initialize from.</param>
        internal override void InitializeFromTransitionTime(TimeZoneInfo.TransitionTime transitionTime)
        {
            throw new ServiceLocalException(Strings.UnsupportedTimeZonePeriodTransitionTarget);
        }

        /// <summary>
        /// Gets the XML element name associated with the transition.
        /// </summary>
        /// <returns>The XML element name associated with the transition.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.AbsoluteDateTransition;
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            bool result = base.TryReadElementFromXml(reader);

            if (!result)
            {
                if (reader.LocalName == XmlElementNames.DateTime)
                {
                    this.dateTime = DateTime.Parse(reader.ReadElementValue(), CultureInfo.InvariantCulture);

                    result = true;
                }
            }

            return result;
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            base.WriteElementsToXml(writer);

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.DateTime,
                this.dateTime);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AbsoluteDateTransition"/> class.
        /// </summary>
        /// <param name="timeZoneDefinition">The time zone definition the transition will belong to.</param>
        internal AbsoluteDateTransition(TimeZoneDefinition timeZoneDefinition)
            : base(timeZoneDefinition)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AbsoluteDateTransition"/> class.
        /// </summary>
        /// <param name="timeZoneDefinition">The time zone definition the transition will belong to.</param>
        /// <param name="targetGroup">The transition group the transition will target.</param>
        internal AbsoluteDateTransition(TimeZoneDefinition timeZoneDefinition, TimeZoneTransitionGroup targetGroup)
            : base(timeZoneDefinition, targetGroup)
        {
        }

        /// <summary>
        /// Gets or sets the absolute date and time when the transition occurs.
        /// </summary>
        internal DateTime DateTime
        {
            get { return this.dateTime; }
            set { this.dateTime = value; }
        }
    }
}
