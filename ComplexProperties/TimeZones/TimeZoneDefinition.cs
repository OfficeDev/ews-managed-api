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
    using System.IO;
    using System.Text;

    /// <summary>
    /// Represents a time zone as defined by the EWS schema.
    /// </summary>
    internal class TimeZoneDefinition : ComplexProperty
    {
        /// <summary>
        /// Prefix for generated ids.
        /// </summary>
        private const string NoIdPrefix = "NoId_";

        private string name;
        private string id;
        private Dictionary<string, TimeZonePeriod> periods = new Dictionary<string, TimeZonePeriod>();
        private Dictionary<string, TimeZoneTransitionGroup> transitionGroups = new Dictionary<string, TimeZoneTransitionGroup>();
        private List<TimeZoneTransition> transitions = new List<TimeZoneTransition>();

        /// <summary>
        /// Compares the transitions.
        /// </summary>
        /// <param name="x">The first transition.</param>
        /// <param name="y">The second transition.</param>
        /// <returns>A negative number if x is less than y, 0 if x and y are equal, a positive number if x is greater than y.</returns>
        private int CompareTransitions(TimeZoneTransition x, TimeZoneTransition y)
        {
            if (x == y)
            {
                return 0;
            }
            else if (x.GetType() == typeof(TimeZoneTransition))
            {
                return -1;
            }
            else if (y.GetType() == typeof(TimeZoneTransition))
            {
                return 1;
            }
            else
            {
                AbsoluteDateTransition firstTransition = (AbsoluteDateTransition)x;
                AbsoluteDateTransition secondTransition = (AbsoluteDateTransition)y;

                return DateTime.Compare(firstTransition.DateTime, secondTransition.DateTime);
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TimeZoneDefinition"/> class.
        /// </summary>
        internal TimeZoneDefinition()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TimeZoneDefinition"/> class.
        /// </summary>
        /// <param name="timeZoneInfo">The time zone info used to initialize this definition.</param>
        internal TimeZoneDefinition(TimeZoneInfo timeZoneInfo)
            : this()
        {
            this.Id = timeZoneInfo.Id;
            this.Name = timeZoneInfo.DisplayName;

            // TimeZoneInfo only supports one standard period, whose bias is the time zone's base
            // offset to UTC.
            TimeZonePeriod standardPeriod = this.CreateStandardPeriod(timeZoneInfo);

            TimeZoneInfo.AdjustmentRule[] adjustmentRules = timeZoneInfo.GetAdjustmentRules();

            if (adjustmentRules.Length == 0)
            {
                // If the time zone info doesn't support Daylight Saving Time, we just need to
                // create one transition to one group with one transition to the standard period.
                this.transitions.Add(new TimeZoneTransition(this,
                    CreateTransitionGroupToPeriod(standardPeriod)));
            }
            else
            {
                for (int i = 0; i < adjustmentRules.Length; i++)
                {
                    TimeZoneTransitionGroup transitionGroup = new TimeZoneTransitionGroup(this, this.transitionGroups.Count.ToString());
                    transitionGroup.InitializeFromAdjustmentRule(timeZoneInfo, adjustmentRules[i]);
                    this.transitionGroups.Add(transitionGroup.Id, transitionGroup);

                    TimeZoneTransition transition;

                    if (i == 0)
                    {
                        // If the first adjustment rule's start date in not undefined (DateTime.MinValue)
                        // we need to add a dummy group with a single, simple transition to the Standard
                        // period and a group containing the transitions mapping to the adjustment rule.
                        if (adjustmentRules[i].DateStart > DateTime.MinValue.Date)
                        {
                            // Add the dummy transition for the standard period.
                            this.transitions.Add(new TimeZoneTransition(this,
                                this.CreateTransitionGroupToPeriod(standardPeriod)));

                            // Add the transition corresponding with the adjustment rule's start date.
                            AbsoluteDateTransition absoluteDateTransition = new AbsoluteDateTransition(this, transitionGroup);
                            absoluteDateTransition.DateTime = adjustmentRules[i].DateStart;

                            transition = absoluteDateTransition;
                        }
                        else
                        {
                            transition = new TimeZoneTransition(this, transitionGroup);
                        }
                    }
                    else
                    {
                        if ((adjustmentRules[i - 1].DateEnd.Year + 1) != adjustmentRules[i].DateStart.Year)
                        {
                            // If the next adjustment rule does not start the year after the previous adjustment rule ends,
                            // we need to add a dummy transition to cover the years in between.
                            AbsoluteDateTransition transitionToDummyGroup = new AbsoluteDateTransition(this,
                                CreateTransitionGroupToPeriod(standardPeriod));
                            transitionToDummyGroup.DateTime = adjustmentRules[i - 1].DateEnd.AddDays(1);

                            this.transitions.Add(transitionToDummyGroup);
                        }

                        AbsoluteDateTransition absoluteDateTransition = new AbsoluteDateTransition(this, transitionGroup);
                        absoluteDateTransition.DateTime = adjustmentRules[i].DateStart;

                        transition = absoluteDateTransition;
                    }

                    this.transitions.Add(transition);
                }

                // If the last adjustment rule's end date is not undefined (DateTime.MaxValue),
                // we need to create another absolute date transition that occurs the date after
                // the last rule's end date. We target this additional transition to a group that
                // contains a single simple transition to the Standard period.
                DateTime lastAdjustmentRuleEndDate = adjustmentRules[adjustmentRules.Length - 1].DateEnd;

                if (lastAdjustmentRuleEndDate < DateTime.MaxValue.Date)
                {
                    AbsoluteDateTransition transitionToDummyGroup = new AbsoluteDateTransition(
                        this,
                        this.CreateTransitionGroupToPeriod(standardPeriod));
                    transitionToDummyGroup.DateTime = lastAdjustmentRuleEndDate.AddDays(1);

                    this.transitions.Add(transitionToDummyGroup);
                }
            }
        }

        /// <summary>
        /// Creates the standard period for the time zone and adds it to the Period collection.
        /// </summary>
        /// <param name="timeZoneInfo">The TimeZoneInfo object whose standard period is to be created.</param>
        /// <returns>The standard period that was created.</returns>
        private TimeZonePeriod CreateStandardPeriod(TimeZoneInfo timeZoneInfo)
        {
            TimeZonePeriod standardPeriod = new TimeZonePeriod();
            standardPeriod.Id = TimeZonePeriod.StandardPeriodId;
            standardPeriod.Name = TimeZonePeriod.StandardPeriodName;
            standardPeriod.Bias = -timeZoneInfo.BaseUtcOffset;

            this.periods.Add(standardPeriod.Id, standardPeriod);

            return standardPeriod;
        }

        /// <summary>
        /// Searches the period collection for a standard period that matches the specifications of the provided adjustment rule.
        /// If one is found, it is returned.
        /// If none are found, a new one is created, added to the collection and returned.
        /// </summary>
        /// <param name="timeZoneInfo">The TimeZoneInfo object that the adjustmentRule applies to.</param>
        /// <param name="adjustmentRule">The AdjustmentRule object that we need a period for.</param>
        /// <returns>The TimeZonePeriod object that was found or created.</returns>
        internal TimeZonePeriod CreateStandardPeriodForAdjustmentRule(TimeZoneInfo timeZoneInfo, TimeZoneInfo.AdjustmentRule adjustmentRule)
        {
            var baseUtcOffsetDelta = adjustmentRule.GetBaseUtcOffsetDelta();

            if (baseUtcOffsetDelta == TimeSpan.Zero)
            {
                // The standard bias of this adjustment rule matches the standard bias for the time zone. Use the standard period.
                return this.periods[TimeZonePeriod.StandardPeriodId];
            }
            else
            {
                TimeSpan periodBias = -timeZoneInfo.BaseUtcOffset - baseUtcOffsetDelta;
                String periodId = String.Format("{0}{1}", TimeZonePeriod.StandardPeriodId, (int)periodBias.TotalMinutes);

                TimeZonePeriod period;
                if (!this.periods.TryGetValue(periodId, out period))
                {
                    period = new TimeZonePeriod();
                    period.Id = periodId;
                    period.Name = TimeZonePeriod.StandardPeriodName;
                    period.Bias = periodBias;

                    this.periods.Add(periodId, period);
                }

                return period;
            }
        }

        /// <summary>
        /// Searches the period collection for a daylight period that matches the specifications of the provided adjustment rule.
        /// If one is found, it is returned.
        /// If none are found, a new one is created, added to the collection and returned.
        /// </summary>
        /// <param name="timeZoneInfo">The TimeZoneInfo object that the adjustmentRule applies to.</param>
        /// <param name="adjustmentRule">The AdjustmentRule object that we need a period for.</param>
        /// <returns>The TimeZonePeriod object that was found or created.</returns>
        internal TimeZonePeriod CreateDaylightPeriodForAdjustmentRule(TimeZoneInfo timeZoneInfo, TimeZoneInfo.AdjustmentRule adjustmentRule)
        {
            var baseUtcOffsetDelta = adjustmentRule.GetBaseUtcOffsetDelta();

            TimeSpan periodBias = -timeZoneInfo.BaseUtcOffset - baseUtcOffsetDelta - adjustmentRule.DaylightDelta;
            String periodId = String.Format("{0}{1}", TimeZonePeriod.DaylightPeriodId, (int)periodBias.TotalMinutes);

            TimeZonePeriod period;
            if (!this.periods.TryGetValue(periodId, out period))
            {
                period = new TimeZonePeriod();
                period.Id = periodId;
                period.Name = TimeZonePeriod.DaylightPeriodName;
                period.Bias = periodBias;

                this.periods.Add(periodId, period);
            }

            return period;
        }

        /// <summary>
        /// Adds a transition group with a single transition to the specified period.
        /// </summary>
        /// <param name="timeZonePeriod">The time zone period.</param>
        /// <returns>A TimeZoneTransitionGroup.</returns>
        private TimeZoneTransitionGroup CreateTransitionGroupToPeriod(TimeZonePeriod timeZonePeriod)
        {
            TimeZoneTransition transitionToPeriod = new TimeZoneTransition(this, timeZonePeriod);

            TimeZoneTransitionGroup transitionGroup = new TimeZoneTransitionGroup(this, this.transitionGroups.Count.ToString());
            transitionGroup.Transitions.Add(transitionToPeriod);

            this.transitionGroups.Add(transitionGroup.Id, transitionGroup);

            return transitionGroup;
        }

        /// <summary>
        /// Reads the attributes from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
        {
            this.name = reader.ReadAttributeValue(XmlAttributeNames.Name);
            this.id = reader.ReadAttributeValue(XmlAttributeNames.Id);

            // EWS can return a TimeZone definition with no Id. Generate a new Id in this case.
            if (string.IsNullOrEmpty(this.id))
            {
                string nameValue = string.IsNullOrEmpty(this.Name) ? string.Empty : this.Name;
                this.Id = NoIdPrefix + Math.Abs(nameValue.GetHashCode()).ToString();
            }
        }

        /// <summary>
        /// Writes the attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            // The Name attribute is only supported in Exchange 2010 and above.
            if (writer.Service.RequestedServerVersion != ExchangeVersion.Exchange2007_SP1)
            {
                writer.WriteAttributeValue(XmlAttributeNames.Name, this.name);
            }

            writer.WriteAttributeValue(XmlAttributeNames.Id, this.id);
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.Periods:
                    do
                    {
                        reader.Read();

                        if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Period))
                        {
                            TimeZonePeriod period = new TimeZonePeriod();
                            period.LoadFromXml(reader);

                            // OM:1648848 Bad timezone data from clients can include duplicate rules
                            // for one year, with duplicate ID. In that case, let the first one win.
                            if (!this.periods.ContainsKey(period.Id))
                            {
                                this.periods.Add(period.Id, period);
                            }
                            else
                            {
                                reader.Service.TraceMessage(
                                    TraceFlags.EwsTimeZones,
                                    string.Format(
                                        "An entry with the same key (Id) '{0}' already exists in Periods. Cannot add another one. Existing entry: [Name='{1}', Bias='{2}']. Entry to skip: [Name='{3}', Bias='{4}'].",
                                        period.Id,
                                        this.Periods[period.Id].Name,
                                        this.Periods[period.Id].Bias,
                                        period.Name,
                                        period.Bias));
                            }
                        }
                    }
                    while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.Periods));

                    return true;
                case XmlElementNames.TransitionsGroups:
                    do
                    {
                        reader.Read();

                        if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.TransitionsGroup))
                        {
                            TimeZoneTransitionGroup transitionGroup = new TimeZoneTransitionGroup(this);

                            transitionGroup.LoadFromXml(reader);

                            this.transitionGroups.Add(transitionGroup.Id, transitionGroup);
                        }
                    }
                    while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.TransitionsGroups));

                    return true;
                case XmlElementNames.Transitions:
                    do
                    {
                        reader.Read();

                        if (reader.IsStartElement())
                        {
                            TimeZoneTransition transition = TimeZoneTransition.Create(this, reader.LocalName);

                            transition.LoadFromXml(reader);

                            this.transitions.Add(transition);
                        }
                    }
                    while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.Transitions));

                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal void LoadFromXml(EwsServiceXmlReader reader)
        {
            this.LoadFromXml(reader, XmlElementNames.TimeZoneDefinition);

            this.transitions.Sort(this.CompareTransitions);
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            // We only emit the full time zone definition against Exchange 2010 servers and above.
            if (writer.Service.RequestedServerVersion != ExchangeVersion.Exchange2007_SP1)
            {
                if (this.periods.Count > 0)
                {
                    writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Periods);

                    foreach (KeyValuePair<string, TimeZonePeriod> keyValuePair in this.periods)
                    {
                        keyValuePair.Value.WriteToXml(writer);
                    }

                    writer.WriteEndElement(); // Periods
                }

                if (this.transitionGroups.Count > 0)
                {
                    writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.TransitionsGroups);

                    foreach (KeyValuePair<string, TimeZoneTransitionGroup> keyValuePair in this.transitionGroups)
                    {
                        keyValuePair.Value.WriteToXml(writer);
                    }

                    writer.WriteEndElement(); // TransitionGroups
                }

                if (this.transitions.Count > 0)
                {
                    writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Transitions);

                    foreach (TimeZoneTransition transition in this.transitions)
                    {
                        transition.WriteToXml(writer);
                    }

                    writer.WriteEndElement(); // Transitions
                }
            }
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal void WriteToXml(EwsServiceXmlWriter writer)
        {
            this.WriteToXml(writer, XmlElementNames.TimeZoneDefinition);
        }

        /// <summary>
        /// Validates this time zone definition.
        /// </summary>
        internal void Validate()
        {
            // The definition must have at least one period, one transition group and one transition,
            // and there must be as many transitions as there are transition groups.
            if (this.periods.Count < 1 || this.transitions.Count < 1 || this.transitionGroups.Count < 1 ||
                this.transitionGroups.Count != this.transitions.Count)
            {
                throw new ServiceLocalException(Strings.InvalidOrUnsupportedTimeZoneDefinition);
            }

            // The first transition must be of type TimeZoneTransition.
            if (this.transitions[0].GetType() != typeof(TimeZoneTransition))
            {
                throw new ServiceLocalException(Strings.InvalidOrUnsupportedTimeZoneDefinition);
            }

            // All transitions must be to transition groups and be either TimeZoneTransition or
            // AbsoluteDateTransition instances.
            foreach (TimeZoneTransition transition in this.transitions)
            {
                Type transitionType = transition.GetType();

                if (transitionType != typeof(TimeZoneTransition) && transitionType != typeof(AbsoluteDateTransition))
                {
                    throw new ServiceLocalException(Strings.InvalidOrUnsupportedTimeZoneDefinition);
                }

                if (transition.TargetGroup == null)
                {
                    throw new ServiceLocalException(Strings.InvalidOrUnsupportedTimeZoneDefinition);
                }
            }

            // All transition groups must be valid.
            foreach (TimeZoneTransitionGroup transitionGroup in this.transitionGroups.Values)
            {
                transitionGroup.Validate();
            }
        }

        /// <summary>
        /// Converts this time zone definition into a TimeZoneInfo structure.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>A TimeZoneInfo representing the same time zone as this definition.</returns>
        internal TimeZoneInfo ToTimeZoneInfo(ExchangeService service)
        {
            this.Validate();

            TimeZoneInfo result;

            // Retrieve the base offset to UTC, standard and daylight display names from
            // the last transition group, which is the one that currently applies given that
            // transitions are ordered chronologically.
            TimeZoneTransitionGroup.CustomTimeZoneCreateParams creationParams =
                this.transitions[this.transitions.Count - 1].TargetGroup.GetCustomTimeZoneCreationParams();

            List<TimeZoneInfo.AdjustmentRule> adjustmentRules = new List<TimeZoneInfo.AdjustmentRule>();

            DateTime startDate = DateTime.MinValue;
            DateTime endDate;
            DateTime effectiveEndDate;

            for (int i = 0; i < this.transitions.Count; i++)
            {
                if (i < this.transitions.Count - 1)
                {
                    endDate = (this.transitions[i + 1] as AbsoluteDateTransition).DateTime;
                    effectiveEndDate = endDate.AddDays(-1);
                }
                else
                {
                    endDate = DateTime.MaxValue;
                    effectiveEndDate = endDate;
                }

                // OM:1648848 Due to bad timezone data from clients the 
                // startDate may not always come before the effectiveEndDate
                if (startDate < effectiveEndDate)
                {
                    TimeZoneInfo.AdjustmentRule adjustmentRule = this.transitions[i].TargetGroup.CreateAdjustmentRule(startDate, effectiveEndDate);

                    if (adjustmentRule != null)
                    {
                        adjustmentRules.Add(adjustmentRule);
                    }

                    startDate = endDate;
                }
                else
                {
                    service.TraceMessage(
                        TraceFlags.EwsTimeZones,
                            string.Format(
                                "The startDate '{0}' is not before the effectiveEndDate '{1}'. Will skip creating adjustment rule.",
                                startDate,
                                effectiveEndDate));
                }
            }

            if (adjustmentRules.Count == 0)
            {
                // If there are no adjustment rule, the time zone does not support Daylight
                // saving time.
                result = TimeZoneInfo.CreateCustomTimeZone(
                    this.Id,
                    creationParams.BaseOffsetToUtc,
                    this.Name,
                    creationParams.StandardDisplayName);
            }
            else
            {
                result = TimeZoneInfo.CreateCustomTimeZone(
                    this.Id,
                    creationParams.BaseOffsetToUtc,
                    this.Name,
                    creationParams.StandardDisplayName,
                    creationParams.DaylightDisplayName,
                    adjustmentRules.ToArray());
            }

            return result;
        }

        /// <summary>
        /// Gets or sets the name of this time zone definition.
        /// </summary>
        internal string Name
        {
            get { return this.name; }
            set { this.name = value; }
        }

        /// <summary>
        /// Gets or sets the Id of this time zone definition.
        /// </summary>
        internal string Id
        {
            get { return this.id; }
            set { this.id = value; }
        }

        /// <summary>
        /// Gets the periods associated with this time zone definition, indexed by Id.
        /// </summary>
        internal Dictionary<string, TimeZonePeriod> Periods
        {
            get { return this.periods; }
        }

        /// <summary>
        /// Gets the transition groups associated with this time zone definition, indexed by Id.
        /// </summary>
        internal Dictionary<string, TimeZoneTransitionGroup> TransitionGroups
        {
            get { return this.transitionGroups; }
        }
    }
}