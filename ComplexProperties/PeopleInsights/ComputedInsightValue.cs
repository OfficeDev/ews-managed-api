// ---------------------------------------------------------------------------
// <copyright file="ComputedInsightValue.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Implements the class for company insight value.</summary>
//-----------------------------------------------------------------------
namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents the ComputedInsightValue.
    /// </summary>
    public sealed class ComputedInsightValue : InsightValue
    {
        /// <summary>
        /// Gets the collection of computed insight
        /// value properties.
        /// </summary>
        public ComputedInsightValuePropertyCollection Properties
        {
            get;
            internal set;
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">XML reader</param>
        /// <returns>Whether the element was read</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.InsightSource:
                    this.InsightSource = reader.ReadElementValue<string>();
                    break;
                case XmlElementNames.Properties:
                    this.Properties = new ComputedInsightValuePropertyCollection();
                    this.Properties.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.Properties);
                    break;
                default:
                    return base.TryReadElementFromXml(reader);
            }

            return true;
        }
    }
}
