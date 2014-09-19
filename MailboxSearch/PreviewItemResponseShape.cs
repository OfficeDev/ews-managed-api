// ---------------------------------------------------------------------------
// <copyright file="PreviewItemResponseShape.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the PreviewItemResponseShape class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents preview item response shape
    /// </summary>
    public sealed class PreviewItemResponseShape
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public PreviewItemResponseShape()
        {
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="baseShape">Preview item base shape</param>
        /// <param name="additionalProperties">Additional properties (must be in form of extended properties)</param>
        public PreviewItemResponseShape(PreviewItemBaseShape baseShape, ExtendedPropertyDefinition[] additionalProperties)
        {
            this.BaseShape = baseShape;
            this.AdditionalProperties = additionalProperties;
        }

        /// <summary>
        /// Mailbox identifier
        /// </summary>
        public PreviewItemBaseShape BaseShape { get; set; }

        /// <summary>
        /// Additional properties (must be in form of extended properties)
        /// </summary>
        public ExtendedPropertyDefinition[] AdditionalProperties { get; set; }
    }
}
