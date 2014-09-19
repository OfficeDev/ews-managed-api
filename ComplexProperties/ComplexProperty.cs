// ---------------------------------------------------------------------------
// <copyright file="ComplexProperty.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ComplexProperty class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// Represents a property that can be sent to or retrieved from EWS.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public abstract class ComplexProperty : ISelfValidate, IJsonSerializable
    {
        private XmlNamespace xmlNamespace = XmlNamespace.Types;

        /// <summary>
        /// Initializes a new instance of the <see cref="ComplexProperty"/> class.
        /// </summary>
        internal ComplexProperty()
        {
        }

        /// <summary>
        /// Gets or sets the namespace.
        /// </summary>
        /// <value>The namespace.</value>
        internal XmlNamespace Namespace
        {
            get { return this.xmlNamespace; }
            set { this.xmlNamespace = value; }
        }

        /// <summary>
        /// Instance was changed.
        /// </summary>
        internal virtual void Changed()
        {
            if (this.OnChange != null)
            {
                this.OnChange(this);
            }
        }

        /// <summary>
        /// Sets value of field.
        /// </summary>
        /// <typeparam name="T">Field type.</typeparam>
        /// <param name="field">The field.</param>
        /// <param name="value">The value.</param>
        internal virtual void SetFieldValue<T>(ref T field, T value)
        {
            bool applyChange;

            if (field == null)
            {
                applyChange = value != null;
            }
            else
            {
                if (field is IComparable)
                {
                    applyChange = (field as IComparable).CompareTo(value) != 0;
                }
                else
                {
                    applyChange = true;
                }
            }

            if (applyChange)
            {
                field = value;
                this.Changed();
            }
        }

        /// <summary>
        /// Clears the change log.
        /// </summary>
        internal virtual void ClearChangeLog()
        {
        }

        /// <summary>
        /// Reads the attributes from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal virtual void ReadAttributesFromXml(EwsServiceXmlReader reader)
        {
        }

        /// <summary>
        /// Reads the text value from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal virtual void ReadTextValueFromXml(EwsServiceXmlReader reader)
        {
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal virtual bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            return false;
        }

        /// <summary>
        /// Tries to read element from XML to patch this property.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal virtual bool TryReadElementFromXmlToPatch(EwsServiceXmlReader reader)
        {
            return false;
        }

        /// <summary>
        /// Writes the attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal virtual void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal virtual void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal virtual void LoadFromXml(
            EwsServiceXmlReader reader,
            XmlNamespace xmlNamespace,
            string xmlElementName)
        {
            this.InternalLoadFromXml(
                reader,
                xmlNamespace,
                xmlElementName,
                this.TryReadElementFromXml);
        }

        /// <summary>
        /// Loads from XML to update itself.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal virtual void UpdateFromXml(
            EwsServiceXmlReader reader,
            XmlNamespace xmlNamespace,
            string xmlElementName)
        {
            this.InternalLoadFromXml(
                reader,
                xmlNamespace,
                xmlElementName,
                this.TryReadElementFromXmlToPatch);
        }

        /// <summary>
        /// Loads from XML
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="readAction"></param>
        private void InternalLoadFromXml(
            EwsServiceXmlReader reader,
            XmlNamespace xmlNamespace,
            string xmlElementName, 
            Func<EwsServiceXmlReader, bool> readAction)
        {
            reader.EnsureCurrentNodeIsStartElement(xmlNamespace, xmlElementName);

            this.ReadAttributesFromXml(reader);

            if (!reader.IsEmptyElement)
            {
                do
                {
                    reader.Read();

                    switch (reader.NodeType)
                    {
                        case XmlNodeType.Element:
                            if (!readAction(reader))
                            {
                                reader.SkipCurrentElement();
                            }
                            break;
                        case XmlNodeType.Text:
                            this.ReadTextValueFromXml(reader);
                            break;
                    }
                }
                while (!reader.IsEndElement(xmlNamespace, xmlElementName));
            }
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal virtual void LoadFromXml(EwsServiceXmlReader reader, string xmlElementName)
        {
            this.LoadFromXml(
                reader,
                this.Namespace,
                xmlElementName);
        }

        /// <summary>
        /// Loads from XML to update this property.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal virtual void UpdateFromXml(EwsServiceXmlReader reader, string xmlElementName)
        {
            this.UpdateFromXml(
                reader,
                this.Namespace,
                xmlElementName);
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service">The service.</param>
        internal virtual void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal virtual void WriteToXml(
            EwsServiceXmlWriter writer,
            XmlNamespace xmlNamespace,
            string xmlElementName)
        {
            writer.WriteStartElement(xmlNamespace, xmlElementName);
            this.WriteAttributesToXml(writer);
            this.WriteElementsToXml(writer);
            writer.WriteEndElement();
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal virtual void WriteToXml(EwsServiceXmlWriter writer, string xmlElementName)
        {
            this.WriteToXml(
                writer,
                this.Namespace,
                xmlElementName);
        }

        /// <summary>
        /// Creates a JSON representation of this object.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        object IJsonSerializable.ToJson(ExchangeService service)
        {
            return this.InternalToJson(service);
        }

        /// <summary>
        /// Serializes the property to a Json value.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        ////internal abstract object InternalToJson(ExchangeService service);
        internal virtual object InternalToJson(ExchangeService service)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Occurs when property changed.
        /// </summary>
        internal event ComplexPropertyChangedDelegate OnChange;

        /// <summary>
        /// Implements ISelfValidate.Validate. Validates this instance.
        /// </summary>
        void ISelfValidate.Validate()
        {
            this.InternalValidate();
        }

        /// <summary>
        ///  Validates this instance.
        /// </summary>
        internal virtual void InternalValidate()
        {
        }
    }
}
