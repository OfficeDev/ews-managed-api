// ---------------------------------------------------------------------------
// <copyright file="AbsoluteDayOfMonthTransition.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AbsoluteDayOfMonthTransition class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a time zone period transition that occurs on a specific day of a specific month.
    /// </summary>
    internal class AbsoluteDayOfMonthTransition : AbsoluteMonthTransition
    {
        private int dayOfMonth;

        /// <summary>
        /// Gets the XML element name associated with the transition.
        /// </summary>
        /// <returns>The XML element name associated with the transition.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.RecurringDateTransition;
        }

        /// <summary>
        /// Creates a timw zone transition time.
        /// </summary>
        /// <returns>A TimeZoneInfo.TransitionTime.</returns>
        internal override TimeZoneInfo.TransitionTime CreateTransitionTime()
        {
            return TimeZoneInfo.TransitionTime.CreateFixedDateRule(
                new DateTime(this.TimeOffset.Ticks),
                this.Month,
                this.DayOfMonth);
        }

        /// <summary>
        /// Initializes this transition based on the specified transition time.
        /// </summary>
        /// <param name="transitionTime">The transition time to initialize from.</param>
        internal override void InitializeFromTransitionTime(TimeZoneInfo.TransitionTime transitionTime)
        {
            base.InitializeFromTransitionTime(transitionTime);

            this.dayOfMonth = transitionTime.Day;
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
                if (reader.LocalName == XmlElementNames.Day)
                {
                    this.dayOfMonth = reader.ReadElementValue<int>();

                    EwsUtilities.Assert(
                        this.dayOfMonth > 0 && this.dayOfMonth <= 31,
                        "AbsoluteDayOfMonthTransition.TryReadElementFromXml",
                        "dayOfMonth is not in the valid 1 - 31 range.");

                    return true;
                }
                else
                {
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
                XmlElementNames.Day,
                this.dayOfMonth);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AbsoluteDayOfMonthTransition"/> class.
        /// </summary>
        /// <param name="timeZoneDefinition">The time zone definition this transition belongs to.</param>
        internal AbsoluteDayOfMonthTransition(TimeZoneDefinition timeZoneDefinition)
            : base(timeZoneDefinition)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AbsoluteDayOfMonthTransition"/> class.
        /// </summary>
        /// <param name="timeZoneDefinition">The time zone definition this transition belongs to.</param>
        /// <param name="targetPeriod">The period the transition will target.</param>
        internal AbsoluteDayOfMonthTransition(TimeZoneDefinition timeZoneDefinition, TimeZonePeriod targetPeriod)
            : base(timeZoneDefinition, targetPeriod)
        {
        }

        /// <summary>
        /// Gets the day of then month when this transition occurs.
        /// </summary>
        internal int DayOfMonth
        {
            get { return this.dayOfMonth; }
        }
    }
}
