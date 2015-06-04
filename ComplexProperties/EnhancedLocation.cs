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
    using System.Text;

    /// <summary>
    /// Represents Enhanced Location.
    /// </summary>
    public sealed class EnhancedLocation : ComplexProperty
    {
        private string displayName;
        private string annotation;
        private PersonaPostalAddress personaPostalAddress;
        
        /// <summary>
        /// Initializes a new instance of the <see cref="EnhancedLocation"/> class.
        /// </summary>
        internal EnhancedLocation()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EnhancedLocation"/> class.
        /// </summary>
        /// <param name="displayName">The location DisplayName.</param>
        public EnhancedLocation(string displayName)
            : this(displayName, String.Empty, new PersonaPostalAddress())
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EnhancedLocation"/> class.
        /// </summary>
        /// <param name="displayName">The location DisplayName.</param>
        /// <param name="annotation">The annotation on the location.</param>
        public EnhancedLocation(string displayName, string annotation)
            : this(displayName, annotation, new PersonaPostalAddress())
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EnhancedLocation"/> class.
        /// </summary>
        /// <param name="displayName">The location DisplayName.</param>
        /// <param name="annotation">The annotation on the location.</param>
        /// <param name="personaPostalAddress">The persona postal address.</param>
        public EnhancedLocation(string displayName, string annotation, PersonaPostalAddress personaPostalAddress)
            : this()
        {
            this.displayName = displayName;
            this.annotation = annotation;
            this.personaPostalAddress = personaPostalAddress;
            this.personaPostalAddress.OnChange += new ComplexPropertyChangedDelegate(PersonaPostalAddress_OnChange);
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
                case XmlElementNames.LocationDisplayName:
                    this.displayName = reader.ReadValue<string>();
                    return true;
                case XmlElementNames.LocationAnnotation:
                    this.annotation = reader.ReadValue<string>();
                    return true;
                case XmlElementNames.PersonaPostalAddress:
                    this.personaPostalAddress = new PersonaPostalAddress();
                    this.personaPostalAddress.LoadFromXml(reader);
                    this.personaPostalAddress.OnChange += new ComplexPropertyChangedDelegate(PersonaPostalAddress_OnChange);
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Gets or sets the Location DisplayName.
        /// </summary>
        public string DisplayName
        {
            get { return this.displayName; }
            set { this.SetFieldValue<string>(ref this.displayName, value); }
        }

        /// <summary>
        /// Gets or sets the Location Annotation.
        /// </summary>
        public string Annotation
        {
            get { return this.annotation; }
            set { this.SetFieldValue<string>(ref this.annotation, value); }
        }

        /// <summary>
        /// Gets or sets the Persona Postal Address.
        /// </summary>
        public PersonaPostalAddress PersonaPostalAddress
        {
            get { return this.personaPostalAddress; }
            set
            {
                if (!this.personaPostalAddress.Equals(value))
                {
                    if (this.personaPostalAddress != null)
                    {
                        this.personaPostalAddress.OnChange -= new ComplexPropertyChangedDelegate(PersonaPostalAddress_OnChange);
                    }

                    this.SetFieldValue<PersonaPostalAddress>(ref this.personaPostalAddress, value);

                    this.personaPostalAddress.OnChange += new ComplexPropertyChangedDelegate(PersonaPostalAddress_OnChange);
                }
            }
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.LocationDisplayName, this.displayName);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.LocationAnnotation, this.annotation);
            this.personaPostalAddress.WriteToXml(writer);
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal override void InternalValidate()
        {
            base.InternalValidate();
            EwsUtilities.ValidateParam(this.displayName, "DisplayName");
            EwsUtilities.ValidateParamAllowNull(this.annotation, "Annotation");
            EwsUtilities.ValidateParamAllowNull(this.personaPostalAddress, "PersonaPostalAddress");
        }

        /// <summary>
        /// PersonaPostalAddress OnChange.
        /// </summary>
        /// <param name="complexProperty">ComplexProperty object.</param>
        private void PersonaPostalAddress_OnChange(ComplexProperty complexProperty)
        {
            this.Changed();
        }
    }
}