// ---------------------------------------------------------------------------
// <copyright file="NoEndRecurrenceRange.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the NoEndRecurrenceRange class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents recurrence range with no end date.
    /// </summary>
    internal sealed class NoEndRecurrenceRange : RecurrenceRange
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="NoEndRecurrenceRange"/> class.
        /// </summary>
        public NoEndRecurrenceRange()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="NoEndRecurrenceRange"/> class.
        /// </summary>
        /// <param name="startDate">The start date.</param>
        public NoEndRecurrenceRange(DateTime startDate)
            : base(startDate)
        {
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <value>The name of the XML element.</value>
        internal override string XmlElementName
        {
            get { return XmlElementNames.NoEndRecurrence; }
        }

        /// <summary>
        /// Setups the recurrence.
        /// </summary>
        /// <param name="recurrence">The recurrence.</param>
        internal override void SetupRecurrence(Recurrence recurrence)
        {
            base.SetupRecurrence(recurrence);

            recurrence.NeverEnds();
        }
    }
}
