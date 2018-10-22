using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Exchange.WebServices.Data;
using System.Collections.Generic;
using System.Diagnostics;

namespace EwsManagedApiTest
{
    /// <summary>
    /// A collection unit tests for testing EWS's Time Zone classes.
    /// </summary>
    [TestClass]
    public class TimeZoneTest
    {
        /// <summary>
        /// An enumeration used to classify a time zone period as either a standard period or a daylight period.
        /// </summary>
        public enum TimeZonePeriodType
        {
            Standard,
            Daylight
        }

        /// <summary>
        /// Verifies that the TimeZoneDefinition contains the period identified by the provided ID string, and that the period
        /// matches the values in the specified TimeZoneInfo object and optionally a specific AdjustmentRule.
        /// </summary>
        /// <param name="tzd">The TimeZoneDefinion object containing the period to validate.</param>
        /// <param name="periodId">The ID of the period to validate.</param>
        /// <param name="tzInfo">The TimeZoneInfo object used to construct the TimeZoneDefinition object.</param>
        /// <param name="adjustmentRule">An optional adjustment rule that applies to the period being validated.</param>
        /// <returns>A enum value indicating whether the period was a Daylight period or a Standard period.</returns>
        private TimeZonePeriodType ValidateTimeZonePeriod(TimeZoneDefinition tzd, String periodId, TimeZoneInfo tzInfo, TimeZoneInfo.AdjustmentRule adjustmentRule)
        {
            TimeZonePeriod period;

            Assert.IsTrue(tzd.Periods.TryGetValue(periodId, out period), "The period was not found in the Time Zone Definition.");

            TimeSpan expectedBias;
            TimeZonePeriodType periodType;

            if (period.IsStandardPeriod)
            {
                expectedBias = tzInfo.BaseUtcOffset;
                if (adjustmentRule != null)
                {
                    expectedBias += adjustmentRule.GetBaseUtcOffsetDelta();
                }

                periodType = TimeZonePeriodType.Standard;
            }
            else
            {
                Assert.IsNotNull(adjustmentRule, "A daylight period was encountered without a matching adjustment rule.");

                expectedBias = tzInfo.BaseUtcOffset + adjustmentRule.GetBaseUtcOffsetDelta() + adjustmentRule.DaylightDelta;

                periodType = TimeZonePeriodType.Daylight;
            }

            Assert.AreEqual(expectedBias, TimeSpan.Zero - period.Bias, "A period bias was not correct.");
            return periodType;
        }

        /// <summary>
        /// Verifies that the TimeZoneDefinition contains the transition group identified by the provided ID string, 
        /// and that the transition group represents a year with no DST period and that the group
        /// matches the values in the specified TimeZoneInfo object and optionally a specific AdjustmentRule.
        /// </summary>
        /// <param name="tzd">The TimeZoneDefinion object containing the transition group to validate.</param>
        /// <param name="groupId">The ID of the transition group to validate.</param>
        /// <param name="tzInfo">The TimeZoneInfo object used to construct the TimeZoneDefinition object.</param>
        /// <param name="adjustmentRule">An optional adjustment rule that applies to the transition group being validated.</param>
        private void ValidateStandardTransitionGroup(TimeZoneDefinition tzd, string groupId, TimeZoneInfo tzInfo, TimeZoneInfo.AdjustmentRule adjustmentRule)
        {
            TimeZoneTransitionGroup transitionGroup;

            Assert.IsTrue(tzd.TransitionGroups.TryGetValue(groupId, out transitionGroup), "The transition group was not found in the time zone definition.");
            if (adjustmentRule != null)
            {
                if (adjustmentRule.DaylightDelta != TimeSpan.Zero)
                {
                    Assert.IsTrue(adjustmentRule.DaylightTransitionStart.HasSameDate(adjustmentRule.DaylightTransitionEnd), "Period transition groups must not be associated with an adjustment rule that has DST enabled unless the DST period starts and ends on the same day.");
                }
            }
            Assert.AreEqual(1, transitionGroup.Transitions.Count, "A period transition group should have only one transition.");

            TimeZoneTransition transition = transitionGroup.Transitions[0];

            Assert.IsNotNull(transition.TargetPeriod, "A transition within a transition group must contain a period.");

            Assert.AreEqual(TimeZonePeriodType.Standard, ValidateTimeZonePeriod(tzd, transition.TargetPeriod.Id, tzInfo, adjustmentRule), "The periods within a period transition must be a standard period.");

            Assert.IsInstanceOfType(transition, typeof(TimeZoneTransition), "The transition within a period transition group must be a standard transition.");
        }

        /// <summary>
        /// Helper function to verify that two TimeZoneTransition structures represent different days of the year.
        /// </summary>
        /// <param name="transition1">The first transition date to verify.</param>
        /// <param name="transition2">The second transition date to verify.</param>
        private void VerifyTransitionsAreOnDifferentDays(TimeZoneTransition transition1, TimeZoneTransition transition2)
        {
            if (transition1.GetType() != transition2.GetType())
            {
                // Assume that if the transitions are different types, they are on different days.
                return;
            }

            if (transition1 is AbsoluteDayOfMonthTransition)
            {
                var absTransition1 = transition1 as AbsoluteDayOfMonthTransition;
                var absTransition2 = transition2 as AbsoluteDayOfMonthTransition;

                Assert.IsFalse((absTransition1.Month == absTransition2.Month) && (absTransition1.DayOfMonth == absTransition2.DayOfMonth), "Daylight start and end date must not be the same.");
            }
            else if (transition1 is RelativeDayOfMonthTransition)
            {
                var relTransition1 = transition1 as RelativeDayOfMonthTransition;
                var relTransition2 = transition2 as RelativeDayOfMonthTransition;

                Assert.IsFalse((relTransition1.Month == relTransition2.Month) 
                    && (relTransition1.WeekIndex == relTransition2.WeekIndex)
                    && (relTransition1.DayOfTheWeek == relTransition2.DayOfTheWeek), 
                    "Daylight start and end date must not be the same.");
            }
            else if (transition1 is AbsoluteDateTransition)
            {
                var absTransition1 = transition1 as AbsoluteDateTransition;
                var absTransition2 = transition2 as AbsoluteDateTransition;

                Assert.IsFalse((absTransition1.DateTime.Year == absTransition2.DateTime.Year)
                    && (absTransition1.DateTime.Month == absTransition2.DateTime.Month) 
                    && (absTransition1.DateTime.Day == absTransition2.DateTime.Day), 
                    "Daylight start and end date must not be the same.");
            }
        }

        /// <summary>
        /// Verifies that the TimeZoneDefinition contains the transition group identified by the provided ID string, 
        /// and that the transition group represents a year with a DST period and that the group
        /// matches the values in the specified TimeZoneInfo object and optionally a specific AdjustmentRule.
        /// </summary>
        /// <param name="tzd">The TimeZoneDefinion object containing the transition group to validate.</param>
        /// <param name="groupId">The ID of the transition group to validate.</param>
        /// <param name="tzInfo">The TimeZoneInfo object used to construct the TimeZoneDefinition object.</param>
        /// <param name="adjustmentRule">The adjustment rule that applies to the transition group being validated.</param>
        private void ValidateDaylightTransitionGroup(TimeZoneDefinition tzd, string groupId, TimeZoneInfo tzInfo, TimeZoneInfo.AdjustmentRule adjustmentRule)
        {
            TimeZoneTransitionGroup transitionGroup;

            Assert.IsTrue(tzd.TransitionGroups.TryGetValue(groupId, out transitionGroup), "The transition group was not found in the time zone definition.");
            Assert.IsNotNull(adjustmentRule, "Transition groups must be associated with an adjustment rule.");
            Assert.AreNotEqual(TimeSpan.Zero, adjustmentRule.DaylightDelta, "Transition groups must be associated with an adjustment rule that has DST enabled.");
            Assert.AreEqual(2, transitionGroup.Transitions.Count, "A transition group should have two transitions.");

            VerifyTransitionsAreOnDifferentDays(transitionGroup.Transitions[0], transitionGroup.Transitions[1]);

            TimeZonePeriodType[] timeZonePeriodTypes = new TimeZonePeriodType[2];

            for (int ix=0; ix<2; ix++)
            {
                TimeZoneTransition transition = transitionGroup.Transitions[ix];

                Assert.IsNotNull(transition.TargetPeriod, "A transition within a transition group must contain a period.");

                timeZonePeriodTypes[ix] = ValidateTimeZonePeriod(tzd, transition.TargetPeriod.Id, tzInfo, adjustmentRule);

                if (transition is AbsoluteDayOfMonthTransition)
                {
                    var absTransition = transition as AbsoluteDayOfMonthTransition;
                    if (timeZonePeriodTypes[ix] == TimeZonePeriodType.Standard)
                    {
                        Assert.IsTrue(adjustmentRule.DaylightTransitionEnd.IsFixedDateRule, "Standard transition should not be a fixed date transition.");
                        Assert.AreEqual((int)adjustmentRule.DaylightTransitionEnd.DayOfWeek, 0, "Standard transition has the wrong DayOfTheWeek.");
                        Assert.AreEqual(adjustmentRule.DaylightTransitionEnd.Month, absTransition.Month, "Standard transition has the wrong DayOfTheWeek.");
                        Assert.AreEqual(adjustmentRule.DaylightTransitionEnd.Day, absTransition.DayOfMonth, "Standard transition has the wrong DayOfMonth.");
                        Assert.AreEqual(adjustmentRule.DaylightTransitionEnd.TimeOfDay.TimeOfDay, absTransition.TimeOffset, "Standard transition has the wrong time offset.");
                    }
                    else
                    {
                        Assert.IsTrue(adjustmentRule.DaylightTransitionStart.IsFixedDateRule, "Daylight transition should not be a fixed date transition.");
                        Assert.AreEqual((int)adjustmentRule.DaylightTransitionStart.DayOfWeek, 0, "Daylight transition has the wrong DayOfTheWeek.");
                        Assert.AreEqual(adjustmentRule.DaylightTransitionStart.Month, absTransition.Month, "Daylight transition has the wrong DayOfTheWeek.");
                        Assert.AreEqual(adjustmentRule.DaylightTransitionStart.Day, absTransition.DayOfMonth, "Daylight transition has the wrong WeekIndex.");
                        Assert.AreEqual(adjustmentRule.DaylightTransitionStart.TimeOfDay.TimeOfDay, absTransition.TimeOffset, "Daylight transition has the wrong time offset.");
                    }
                }
                else if (transition is RelativeDayOfMonthTransition)
                {
                    var relTransition = transition as RelativeDayOfMonthTransition;
                    if (timeZonePeriodTypes[ix] == TimeZonePeriodType.Standard)
                    {
                        Assert.IsFalse(adjustmentRule.DaylightTransitionEnd.IsFixedDateRule, "Standard transition should be a fixed date transition.");
                        Assert.AreEqual((int)adjustmentRule.DaylightTransitionEnd.DayOfWeek, (int)relTransition.DayOfTheWeek, "Standard transition has the wrong DayOfTheWeek.");
                        Assert.AreEqual(adjustmentRule.DaylightTransitionEnd.Month, relTransition.Month, "Standard transition has the wrong DayOfTheWeek.");
                        Assert.AreEqual((adjustmentRule.DaylightTransitionEnd.Week == 5) ? -1 : adjustmentRule.DaylightTransitionEnd.Week, relTransition.WeekIndex, "Standard transition has the wrong WeekIndex.");
                        Assert.AreEqual(adjustmentRule.DaylightTransitionEnd.TimeOfDay.TimeOfDay, relTransition.TimeOffset, "Standard transition has the wrong time offset.");
                    }
                    else
                    {
                        Assert.IsFalse(adjustmentRule.DaylightTransitionStart.IsFixedDateRule, "Daylight transition should be a fixed date transition.");
                        Assert.AreEqual((int)adjustmentRule.DaylightTransitionStart.DayOfWeek, (int)relTransition.DayOfTheWeek, "Daylight transition has the wrong DayOfTheWeek.");
                        Assert.AreEqual(adjustmentRule.DaylightTransitionStart.Month, relTransition.Month, "Daylight transition has the wrong DayOfTheWeek.");
                        Assert.AreEqual((adjustmentRule.DaylightTransitionStart.Week == 5) ? -1 : adjustmentRule.DaylightTransitionStart.Week, relTransition.WeekIndex, "Daylight transition has the wrong WeekIndex.");
                        Assert.AreEqual(adjustmentRule.DaylightTransitionStart.TimeOfDay.TimeOfDay, relTransition.TimeOffset, "Daylight transition has the wrong time offset.");
                    }
                }
                else if (transition is AbsoluteDateTransition)
                {
                    var absTransition = transition as AbsoluteDateTransition;
                    if (timeZonePeriodTypes[ix] == TimeZonePeriodType.Standard)
                    {
                        Assert.IsTrue(adjustmentRule.DaylightTransitionEnd.IsFixedDateRule, "Standard transition should not be a fixed date transition.");
                        Assert.AreEqual((int)adjustmentRule.DaylightTransitionEnd.DayOfWeek, 0, "Standard transition has the wrong DayOfTheWeek.");
                        Assert.AreEqual(adjustmentRule.DaylightTransitionEnd.Month, absTransition.DateTime.Month, "Standard transition has the wrong DayOfTheWeek.");
                        Assert.AreEqual(adjustmentRule.DaylightTransitionEnd.Day, absTransition.DateTime.Day, "Standard transition has the wrong DayOfMonth.");
                        Assert.IsTrue(absTransition.DateTime.Year > 0, "Standard transition year is not set.");
                        Assert.AreEqual(adjustmentRule.DaylightTransitionEnd.TimeOfDay.TimeOfDay, absTransition.DateTime.TimeOfDay, "Standard transition has the wrong time offset.");
                    }
                    else
                    {
                        Assert.IsTrue(adjustmentRule.DaylightTransitionStart.IsFixedDateRule, "Daylight transition should not be a fixed date transition.");
                        Assert.AreEqual((int)adjustmentRule.DaylightTransitionStart.DayOfWeek, 0, "Daylight transition has the wrong DayOfTheWeek.");
                        Assert.AreEqual(adjustmentRule.DaylightTransitionStart.Month, absTransition.DateTime.Month, "Daylight transition has the wrong DayOfTheWeek.");
                        Assert.AreEqual(adjustmentRule.DaylightTransitionStart.Day, absTransition.DateTime.Day, "Daylight transition has the wrong WeekIndex.");
                        Assert.IsTrue(absTransition.DateTime.Year > 0, "Daylight transition year is not set.");
                        Assert.AreEqual(adjustmentRule.DaylightTransitionStart.TimeOfDay.TimeOfDay, absTransition.DateTime.TimeOfDay, "Daylight transition has the wrong time offset.");
                    }
                }
                else 
                {
                    Assert.Fail("Standard transitions are not allowed in a transition group.");
                }
            }

            Assert.IsTrue(timeZonePeriodTypes[0] != timeZonePeriodTypes[1], "A time zone transition group must contain one standard period and one daylight period.");
        }

        /// <summary>
        /// A helper function that will create a TimeZoneDefinition object from the provided TimeZoneInfo object and
        /// verify that the Dynamic DST rules represented by the created TimeZoneDefinition object matches those
        /// of the TimeZoneInfo object.
        /// </summary>
        /// <param name="timeZoneInfo">The TimeZoneInfo object representing the time zone to be tested.</param>
        public void TestTimeZone(TimeZoneInfo timeZoneInfo)
        {
            TimeZoneDefinition tzd = new TimeZoneDefinition(timeZoneInfo);
            var transitions = tzd.GetTransitions();
            if ((transitions == null) || (transitions.Count < 1))
            {
                Assert.Fail("No transitions found.");
            }

            var adjustmentRules = timeZoneInfo.GetAdjustmentRules();
            if (adjustmentRules.Length < 1)
            {
                adjustmentRules = null;
            }
            int adjustmentRuleIndex = -1;
            TimeZoneInfo.AdjustmentRule lastAdjustmentRule = null;

            foreach (var transition in transitions)
            {
                Assert.IsNull(transition.TargetPeriod, "Time zone transitions cannot reference time zone periods.");
                Assert.IsNotNull(transition.TargetGroup, "Time zone transitions must referenca a time zone group.");

                TimeZoneInfo.AdjustmentRule adjustmentRule = null;
                if (transition.GetType() == typeof(TimeZoneTransition))
                {
                    if (adjustmentRules != null)
                    {
                        // If there are adjustment rules, the first rule will apply to the transition if the date on the rule is MinDate.Date
                        if (adjustmentRules[0].DateStart == DateTime.MinValue.Date)
                        {
                            adjustmentRuleIndex = 0;
                            adjustmentRule = adjustmentRules[adjustmentRuleIndex];
                        }
                    }
                }
                else if (transition.GetType() == typeof(AbsoluteDateTransition))
                {
                    AbsoluteDateTransition absTransition = (AbsoluteDateTransition)transition;

                    Assert.IsNotNull(adjustmentRules, "The time zone definition cannot have any absolute date transitions if the time zone does not have any adjustment rules.");

                    if (lastAdjustmentRule != null)
                    {
                        Assert.AreEqual(lastAdjustmentRule.DateEnd.AddDays(1), absTransition.DateTime, "The transition date should be the next day after the end of the previous adjustment rule.");

                        if (adjustmentRuleIndex < (adjustmentRules.Length - 1))
                        {
                            if (absTransition.DateTime < adjustmentRules[adjustmentRuleIndex + 1].DateStart)
                            {
                                // We have a hole in between the adjustment rules.
                                adjustmentRule = null;
                            }
                            else if (absTransition.DateTime == adjustmentRules[adjustmentRuleIndex + 1].DateStart)
                            {
                                adjustmentRule = adjustmentRules[++adjustmentRuleIndex];
                            }
                            else
                            {
                                // Should not be possible, but just in case.
                                Assert.Fail("The transition start date cannot be greater than the next available adjustment rule.");
                            }
                        }
                        else
                        {
                            adjustmentRule = null;
                        }
                    }
                    else
                    {
                        if (adjustmentRuleIndex < (adjustmentRules.Length - 1))
                        {
                            if (absTransition.DateTime < adjustmentRules[adjustmentRuleIndex + 1].DateStart)
                            {
                                Assert.Fail("A transition found with a start date that is less than the next available adjustment rule.");
                            }
                            else if (absTransition.DateTime == adjustmentRules[adjustmentRuleIndex + 1].DateStart)
                            {
                                adjustmentRule = adjustmentRules[++adjustmentRuleIndex];
                            }
                            else
                            {
                                // Should not be possible, but just in case.
                                Assert.Fail("The transition start date cannot be greater than the next available adjustment rule.");
                            }
                        }
                        else
                        {
                            Assert.Fail("A transition found with no more adjustment rules to pick from.");
                        }
                    }
                }
                else
                {
                    Assert.Fail("Unexpected transition type in the transition collection.");
                }

                if (transition.TargetGroup.SupportsDaylight)
                {
                    ValidateDaylightTransitionGroup(tzd, transition.TargetGroup.Id, timeZoneInfo, adjustmentRule);
                }
                else
                {
                    ValidateStandardTransitionGroup(tzd, transition.TargetGroup.Id, timeZoneInfo, adjustmentRule);
                }

                lastAdjustmentRule = adjustmentRule;
            }

            if (adjustmentRules != null)
            {
                Assert.AreEqual(adjustmentRules.Length - 1, adjustmentRuleIndex, "There are unprocessed adjustment rules in the time zone.");
            }

            if (lastAdjustmentRule != null)
            {
                // If the last transition corresponds with an adjustment rule, the adjustment rule should terminate with the Max date value.
                Assert.AreEqual(DateTime.MaxValue.Date, lastAdjustmentRule.DateEnd, "The last adjustment rule does not end with the max date value. An additional transition should have been created.");
            }
        }

        /// <summary>
        /// A unit test that tests the creation of a TimeZoneDefinion object for every time zone known to the .Net runtime.
        /// </summary>
        [TestMethod]
        public void TimeZoneTest_TimeZoneDefinitions()
        {
            int failureCount = 0;
            foreach (TimeZoneInfo timeZoneInfo in TimeZoneInfo.GetSystemTimeZones())
            {
                try
                {
                    TestTimeZone(timeZoneInfo);
                }
                catch(Exception err)
                {
                    //Assert.Fail("The time zone '{0}' failed with error [{1}].", timeZoneInfo.Id, err.Message);
                    Debug.Print("The time zone '{0}' failed with error [{1}].", timeZoneInfo.Id, err.Message);
                    failureCount++;
                }
            }

            Assert.AreEqual(0, failureCount, "One or more time zones failed the TimeZoneDefinition test.");
        }

        /// <summary>
        /// This test is useful for when you need to step through the creation of a specific time zone definition.
        /// It is here for debugging purposes only, which is why it has the [Ignore] attribute.
        /// </summary>
        [Ignore]
        [TestMethod]
        public void TimeZoneTest_TimeZoneDefinition()
        {
            TestTimeZone(TimeZoneInfo.GetSystemTimeZones().Single(x => x.Id == "Venezuela Standard Time"));
        }
    }
}
