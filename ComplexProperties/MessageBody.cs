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
// <summary>Defines the MessageBody class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents the body of a message.
    /// </summary>
    public class MessageBody : ComplexProperty
    {
        private BodyType bodyType;
        private string text;

        /// <summary>
        /// Initializes a new instance of the <see cref="MessageBody"/> class.
        /// </summary>
        public MessageBody()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="MessageBody"/> class.
        /// </summary>
        /// <param name="bodyType">The type of the message body's text.</param>
        /// <param name="text">The text of the message body.</param>
        public MessageBody(BodyType bodyType, string text)
            : this()
        {
            this.bodyType = bodyType;
            this.text = text;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="MessageBody"/> class.
        /// </summary>
        /// <param name="text">The text of the message body, assumed to be HTML.</param>
        public MessageBody(string text)
            : this(BodyType.HTML, text)
        {
        }

        /// <summary>
        /// Defines an implicit conversation between a string and MessageBody.
        /// </summary>
        /// <param name="textBody">The string to convert to MessageBody, assumed to be HTML.</param>
        /// <returns>A MessageBody initialized with the specified string.</returns>
        public static implicit operator MessageBody(string textBody)
        {
            return new MessageBody(BodyType.HTML, textBody);
        }

        /// <summary>
        /// Defines an implicit conversion of MessageBody into a string.
        /// </summary>
        /// <param name="messageBody">The MessageBody to convert to a string.</param>
        /// <returns>A string containing the text of the MessageBody.</returns>
        public static implicit operator string(MessageBody messageBody)
        {
            EwsUtilities.ValidateParam(messageBody, "messageBody");

            return messageBody.Text;
        }

        /// <summary>
        /// Reads attributes from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
        {
            this.bodyType = reader.ReadAttributeValue<BodyType>(XmlAttributeNames.BodyType);
        }

        /// <summary>
        /// Reads text value from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadTextValueFromXml(EwsServiceXmlReader reader)
        {
            this.text = reader.ReadValue();
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service">The service.</param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            foreach (string key in jsonProperty.Keys)
            {
                switch (key)
                {
                    case XmlAttributeNames.BodyType:
                        this.bodyType = jsonProperty.ReadEnumValue<BodyType>(key);
                        break;
                    case JsonObject.JsonValueString:
                        this.text = jsonProperty.ReadAsString(key);
                        break;
                    default:
                        break; 
                }
            }
        }

        /// <summary>
        /// Writes attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteAttributeValue(XmlAttributeNames.BodyType, this.BodyType);
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            if (!string.IsNullOrEmpty(this.Text))
            {
                writer.WriteValue(this.Text, XmlElementNames.Body);
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
            JsonObject jsonProperty = new JsonObject();

            jsonProperty.Add(XmlAttributeNames.BodyType, this.BodyType);

            if (!string.IsNullOrEmpty(this.Text))
            {
                jsonProperty.Add(JsonObject.JsonValueString, this.Text);
            }
            return jsonProperty;
        }

        /// <summary>
        /// Gets or sets the type of the message body's text.
        /// </summary>
        public BodyType BodyType
        {
            get { return this.bodyType; }
            set { this.SetFieldValue<BodyType>(ref this.bodyType, value); }
        }

        /// <summary>
        /// Gets or sets the text of the message body.
        /// </summary>
        public string Text
        {
            get { return this.text; }
            set { this.SetFieldValue<string>(ref this.text, value); }
        }

        #region Object method overrides
        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override string ToString()
        {
            return (this.Text == null) ? string.Empty : this.Text;
        }
        #endregion
    }
}
