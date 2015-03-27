// ---------------------------------------------------------------------------
// <copyright file="TimeZoneTransitionGroup.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the TimeZoneTransitionGroup class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a group of time zone period transitions.
    /// </summary>
    internal class TimeZoneTransitionGroup : ComplexProperty
    {
        private TimeZoneDefinition timeZoneDefinition;
        private string id;
        private List<TimeZoneTransition> transitions = new List<TimeZoneTransition>();
        private TimeZoneTransition transitionToStandard;
        private TimeZoneTransition transitionToDaylight;

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal void LoadFromXml(EwsServiceXmlReader reader)
        {
            this.LoadFromXml(reader, XmlElementNames.TransitionsGroup);
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal void WriteToXml(EwsServiceXmlWriter writer)
        {
            this.WriteToXml(writer, XmlElementNames.TransitionsGroup);
        }

        /// <summary>
        /// Reads the attributes from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
        {
            this.id = reader.ReadAttributeValue(XmlAttributeNames.Id);
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service">The service.</param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            base.LoadFromJson(jsonProperty, service);

            foreach (string key in jsonProperty.Keys)
            {
                switch (key)
                {
                    case XmlAttributeNames.Id:
                        this.id = jsonProperty.ReadAsString(key);
                        break;

                    case XmlElementNames.Transition:

                        foreach (object uncastJsonTransition in jsonProperty.ReadAsArray(key))
                        {
                            JsonObject jsonTransition = uncastJsonTransition as JsonObject;
                            TimeZoneTransition transition = TimeZoneTransition.Create(this.timeZoneDefinition, jsonTransition.ReadTypeString());

                            transition.LoadFromJson(jsonTransition, service);

                            this.transitions.Add(transition);
                        }
                        break;
                }
            }
        }

        /// <summary>
        /// Writes the attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteAttributeValue(XmlAttributeNames.Id, this.id);
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            reader.EnsureCurrentNodeIsStartElement();

            TimeZoneTransition transition = TimeZoneTransition.Create(this.timeZoneDefinition, reader.LocalName);

            transition.LoadFromXml(reader);

            EwsUtilities.Assert(
                transition.TargetPeriod != null,
                "TimeZoneTransitionGroup.TryReadElementFromXml",
                "The transition's target period is null.");

            this.transitions.Add(transition);

            return true;
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            foreach (TimeZoneTransition transition in this.transitions)
            {
                transition.WriteToXml(writer);
            }
        }

        /// <summary>
        /// Serializes the property to a Json value.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        internal override object InternalToJson(ExchangeService service)
        {
            JsonObject jsonTimeZoneTransitionGroup = new JsonObject();

            jsonTimeZoneTransitionGroup.Add(XmlAttributeNames.Id, this.id);

            List<object> jsonTransitions = new List<object>();
            foreach (TimeZoneTransition transition in this.transitions)
            {
                jsonTransitions.Add(transition.InternalToJson(service));
            }

            jsonTimeZoneTransitionGroup.Add(XmlElementNames.Transitions, jsonTransitions.ToArray());

            return jsonTimeZoneTransitionGroup;
        }

        /// <summary>
        /// Initializes this transition group based on the specified asjustment rule.
        /// </summary>
        /// <param name="adjustmentRule">The adjustment rule to initialize from.</param>
        /// <param name="standardPeriod">A reference to the pre-created standard period.</param>
        internal virtual void InitializeFromAdjustmentRule(TimeZoneInfo.AdjustmentRule adjustmentRule, TimeZonePeriod standardPeriod)
        {
            TimeZonePeriod daylightPeriod = new TimeZonePeriod();

            // Generate an Id of the form "Daylight/2008"
            daylightPeriod.Id = string.Format(
                "{0}/{1}",
                TimeZonePeriod.DaylightPeriodId,
                adjustmentRule.DateStart.Year);
            daylightPeriod.Name = TimeZonePeriod.DaylightPeriodName;
            daylightPeriod.Bias = standardPeriod.Bias - adjustmentRule.DaylightDelta;

            this.timeZoneDefinition.Periods.Add(daylightPeriod.Id, daylightPeriod);

            this.transitionToDaylight = TimeZoneTransition.CreateTimeZoneTransition(
                this.timeZoneDefinition,
                daylightPeriod,
                adjustmentRule.DaylightTransitionStart);

            TimeZonePeriod standardPeriodToSet = new TimeZonePeriod();
            standardPeriodToSet.Id = string.Format(
                "{0}/{1}",
                standardPeriod.Id,
                adjustmentRule.DateStart.Year);
            standardPeriodToSet.Name = standardPeriod.Name;
            standardPeriodToSet.Bias = standardPeriod.Bias;
            this.timeZoneDefinition.Periods.Add(standardPeriodToSet.Id, standardPeriodToSet);

            this.transitionToStandard = TimeZoneTransition.CreateTimeZoneTransition(
                this.timeZoneDefinition,
                standardPeriodToSet,
                adjustmentRule.DaylightTransitionEnd);

            this.transitions.Add(this.transitionToDaylight);
            this.transitions.Add(this.transitionToStandard);
        }

        /// <summary>
        /// Validates this transition group.
        /// </summary>
        internal void Validate()
        {
            // There must be exactly one or two transitions in the group.
            if (this.transitions.Count < 1 || this.transitions.Count > 2)
            {
                throw new ServiceLocalException(Strings.InvalidOrUnsupportedTimeZoneDefinition);
            }

            // If there is only one transition, it must be of type TimeZoneTransition
            if (this.transitions.Count == 1 && !(this.transitions[0].GetType() == typeof(TimeZoneTransition)))
            {
                throw new ServiceLocalException(Strings.InvalidOrUnsupportedTimeZoneDefinition);
            }

            // If there are two transitions, none of them should be of type TimeZoneTransition
            if (this.transitions.Count == 2)
            {
                foreach (TimeZoneTransition transition in this.transitions)
                {
                    if (transition.GetType() == typeof(TimeZoneTransition))
                    {
                        throw new ServiceLocalException(Strings.InvalidOrUnsupportedTimeZoneDefinition);
                    }
                }
            }

            // All the transitions in the group must be to a period.
            foreach (TimeZoneTransition transition in this.transitions)
            {
                if (transition.TargetPeriod == null)
                {
                    throw new ServiceLocalException(Strings.InvalidOrUnsupportedTimeZoneDefinition);
                }
            }
        }

        /// <summary>
        /// Represents custom time zone creation parameters.
        /// </summary>
        internal class CustomTimeZoneCreateParams
        {
            private TimeSpan baseOffsetToUtc;
            private string standardDisplayName;
            private string daylightDisplayName;

            /// <summary>
            /// Initializes a new instance of the <see cref="CustomTimeZoneCreateParams"/> class.
            /// </summary>
            internal CustomTimeZoneCreateParams()
            {
            }

            /// <summary>
            /// Gets or sets the base offset to UTC.
            /// </summary>
            internal TimeSpan BaseOffsetToUtc
            {
                get { return this.baseOffsetToUtc; }
                set { this.baseOffsetToUtc = value; }
            }

            /// <summary>
            /// Gets or sets the display name of the standard period.
            /// </summary>
            internal string StandardDisplayName
            {
                get { return this.standardDisplayName; }
                set { this.standardDisplayName = value; }
            }

            /// <summary>
            /// Gets or sets the display name of the daylight period.
            /// </summary>
            internal string DaylightDisplayName
            {
                get { return this.daylightDisplayName; }
                set { this.daylightDisplayName = value; }
            }

            /// <summary>
            /// Gets a value indicating whether the custom time zone should have a daylight period.
            /// </summary>
            /// <value>
            ///     <c>true</c> if the custom time zone should have a daylight period; otherwise, <c>false</c>.
            /// </value>
            internal bool HasDaylightPeriod
            {
                get { return !string.IsNullOrEmpty(this.daylightDisplayName); }
            }
        }

        /// <summary>
        /// Gets a value indicating whether this group contains a transition to the Daylight period.
        /// </summary>
        /// <value><c>true</c> if this group contains a transition to daylight; otherwise, <c>false</c>.</value>
        internal bool SupportsDaylight
        {
            get { return this.transitions.Count == 2; }
        }

        /// <summary>
        /// Initializes the private members holding references to the transitions to the Daylight
        /// and Standard periods.
        /// </summary>
        private void InitializeTransitions()
        {
            if (this.transitionToStandard == null)
            {
                foreach (TimeZoneTransition transition in this.transitions)
                {
                    if (transition.TargetPeriod.IsStandardPeriod || (this.transitions.Count == 1))
                    {
                        this.transitionToStandard = transition;
                    }
                    else
                    {
                        this.transitionToDaylight = transition;
                    }
                }
            }

            // If we didn't find a Standard period, this is an invalid time zone group.
            if (this.transitionToStandard == null)
            {
                throw new ServiceLocalException(Strings.InvalidOrUnsupportedTimeZoneDefinition);
            }
        }

        /// <summary>
        /// Gets the transition to the Daylight period.
        /// </summary>
        private TimeZoneTransition TransitionToDaylight
        {
            get
            {
                this.InitializeTransitions();

                return this.transitionToDaylight;
            }
        }

        /// <summary>
        /// Gets the transition to the Standard period.
        /// </summary>
        private TimeZoneTransition TransitionToStandard
        {
            get
            {
                this.InitializeTransitions();

                return this.transitionToStandard;
            }
        }

        /// <summary>
        /// Gets the offset to UTC based on this group's transitions.
        /// </summary>
        internal CustomTimeZoneCreateParams GetCustomTimeZoneCreationParams()
        {
            CustomTimeZoneCreateParams result = new CustomTimeZoneCreateParams();

            if (this.TransitionToDaylight != null)
            {
                result.DaylightDisplayName = this.TransitionToDaylight.TargetPeriod.Name;
            }

            result.StandardDisplayName = this.TransitionToStandard.TargetPeriod.Name;

            // Assume that the standard period's offset is the base offset to UTC.
            // EWS returns a positive offset for time zones that are behind UTC, and
            // a negative one for time zones ahead of UTC. TimeZoneInfo does it the other
            // way around.
            result.BaseOffsetToUtc = -this.TransitionToStandard.TargetPeriod.Bias;

            return result;
        }

        /// <summary>
        /// Gets the delta offset for the daylight.
        /// </summary>
        /// <returns></returns>
        internal TimeSpan GetDaylightDelta()
        {
            if (this.SupportsDaylight)
            {
                // EWS returns a positive offset for time zones that are behind UTC, and
                // a negative one for time zones ahead of UTC. TimeZoneInfo does it the other
                // way around.
                return this.TransitionToStandard.TargetPeriod.Bias - this.TransitionToDaylight.TargetPeriod.Bias;
            }
            else
            {
                return TimeSpan.Zero;
            }
        }

        /// <summary>
        /// Creates a time zone adjustment rule.
        /// </summary>
        /// <param name="startDate">The start date of the adjustment rule.</param>
        /// <param name="endDate">The end date of the adjustment rule.</param>
        /// <returns>An TimeZoneInfo.AdjustmentRule.</returns>
        internal TimeZoneInfo.AdjustmentRule CreateAdjustmentRule(DateTime startDate, DateTime endDate)
        {
            // If there is only one transition, we can't create an adjustment rule. We have to assume
            // that the base offset to UTC is unchanged.
            if (this.transitions.Count == 1)
            {
                return null;
            }

            return TimeZoneInfo.AdjustmentRule.CreateAdjustmentRule(
                startDate.Date,
                endDate.Date,
                this.GetDaylightDelta(),
                this.TransitionToDaylight.CreateTransitionTime(),
                this.TransitionToStandard.CreateTransitionTime());
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TimeZoneTransitionGroup"/> class.
        /// </summary>
        /// <param name="timeZoneDefinition">The time zone definition.</param>
        internal TimeZoneTransitionGroup(TimeZoneDefinition timeZoneDefinition)
            : base()
        {
            this.timeZoneDefinition = timeZoneDefinition;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TimeZoneTransitionGroup"/> class.
        /// </summary>
        /// <param name="timeZoneDefinition">The time zone definition.</param>
        /// <param name="id">The Id of the new transition group.</param>
        internal TimeZoneTransitionGroup(TimeZoneDefinition timeZoneDefinition, string id)
            : this(timeZoneDefinition)
        {
            this.id = id;
        }

        /// <summary>
        /// Gets or sets the id of this group.
        /// </summary>
        internal string Id
        {
            get { return this.id; }
            set { this.id = value; }
        }

        /// <summary>
        /// Gets the transitions in this group.
        /// </summary>
        internal List<TimeZoneTransition> Transitions
        {
            get { return this.transitions; }
        }
    }
}
