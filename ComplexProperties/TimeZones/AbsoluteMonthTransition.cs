// ---------------------------------------------------------------------------
// <copyright file="AbsoluteMonthTransition.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AbsoluteMonthTransition class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the base class for all recurring time zone period transitions.
    /// </summary>
    internal abstract class AbsoluteMonthTransition : TimeZoneTransition
    {
        private TimeSpan timeOffset;
        private int month;

        /// <summary>
        /// Initializes this transition based on the specified transition time.
        /// </summary>
        /// <param name="transitionTime">The transition time to initialize from.</param>
        internal override void InitializeFromTransitionTime(TimeZoneInfo.TransitionTime transitionTime)
        {
            base.InitializeFromTransitionTime(transitionTime);

            this.timeOffset = transitionTime.TimeOfDay.TimeOfDay;
            this.month = transitionTime.Month;
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            if (base.TryReadElementFromXml(reader))
            {
                return true;
            }
            else
            {
                switch (reader.LocalName)
                {
                    case XmlElementNames.TimeOffset:
                        this.timeOffset = EwsUtilities.XSDurationToTimeSpan(reader.ReadElementValue());
                        return true;
                    case XmlElementNames.Month:
                        this.month = reader.ReadElementValue<int>();

                        EwsUtilities.Assert(
                            this.month > 0 && this.month <= 12,
                            "AbsoluteMonthTransition.TryReadElementFromXml",
                            "month is not in the valid 1 - 12 range.");

                        return true;
                    default:
                        return false;
                }
            }
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
                XmlElementNames.TimeOffset,
                EwsUtilities.TimeSpanToXSDuration(this.timeOffset));

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.Month,
                this.month);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AbsoluteMonthTransition"/> class.
        /// </summary>
        /// <param name="timeZoneDefinition">The time zone definition this transition belongs to.</param>
        internal AbsoluteMonthTransition(TimeZoneDefinition timeZoneDefinition)
            : base(timeZoneDefinition)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AbsoluteMonthTransition"/> class.
        /// </summary>
        /// <param name="timeZoneDefinition">The time zone definition this transition belongs to.</param>
        /// <param name="targetPeriod">The period the transition will target.</param>
        internal AbsoluteMonthTransition(TimeZoneDefinition timeZoneDefinition, TimeZonePeriod targetPeriod)
            : base(timeZoneDefinition, targetPeriod)
        {
        }

        /// <summary>
        /// Gets the time offset from midnight when the transition occurs.
        /// </summary>
        internal TimeSpan TimeOffset
        {
            get { return this.timeOffset; }
        }

        /// <summary>
        /// Gets the month when the transition occurs.
        /// </summary>
        internal int Month
        {
            get { return this.month; }
        }
    }
}
