#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

//-----------------------------------------------------------------------
// <summary>Defines the GroupMemberPropertyDefinition class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents the definition of the GroupMember property.
    /// </summary>
    internal sealed class GroupMemberPropertyDefinition : ServiceObjectPropertyDefinition
    {
        /// <summary>
        /// FieldUri of IndexedFieldURI for a group member.
        /// </summary>
        private const string FieldUri = "distributionlist:Members:Member";

        /// <summary>
        /// Member key.
        /// Maps to the Index attribute of IndexedFieldURI element.
        /// </summary>
        private string key;

        /// <summary>
        /// Initializes a new instance of the <see cref="GroupMemberPropertyDefinition"/> class.
        /// </summary>
        /// <param name="key">The member's key.</param>
        public GroupMemberPropertyDefinition(string key)
            : base(FieldUri)
        {
            this.key = key;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GroupMemberPropertyDefinition"/> class without key.
        /// </summary>
        internal GroupMemberPropertyDefinition()
            : base(FieldUri)
        {
        }

        /// <summary>
        /// Gets or sets the member's key.
        /// </summary>
        public string Key
        {
            get
            {
                return this.key;
            }

            set
            {
                this.key = value;
            }
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.IndexedFieldURI;
        }

        protected override string GetJsonType()
        {
            return JsonNames.PathToIndexedFieldType;
        }

        /// <summary>
        /// Writes the attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            base.WriteAttributesToXml(writer);
            writer.WriteAttributeValue(XmlAttributeNames.FieldIndex, this.Key);
        }

        /// <summary>
        /// Adds the json properties.
        /// </summary>
        /// <param name="jsonPropertyDefinition">The json property definition.</param>
        /// <param name="service">The service.</param>
        internal override void AddJsonProperties(JsonObject jsonPropertyDefinition, ExchangeService service)
        {
            base.AddJsonProperties(jsonPropertyDefinition, service);
            jsonPropertyDefinition.Add(XmlAttributeNames.FieldIndex, this.Key);
        }

        /// <summary>
        /// Gets the property definition's printable name.
        /// </summary>
        /// <returns>
        /// The property definition's printable name.
        /// </returns>
        internal override string GetPrintableName()
        {
            return string.Format("{0}:{1}", FieldUri, this.Key);
        }

        /// <summary>
        /// Gets the property type.
        /// </summary>
        public override Type Type
        {
            get { return typeof(string); }
        }
    }
}
