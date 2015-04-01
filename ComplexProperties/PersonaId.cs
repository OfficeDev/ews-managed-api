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

    /// <summary>
    /// Represents the Id of a Persona.
    /// </summary>
    public sealed class PersonaId : ServiceId
    {
        /// <summary>
        /// Creates a new instance of the <see cref="PersonaId"/> class.
        /// </summary>
        internal PersonaId()
            : base()
        {
        }

        /// <summary>
        /// Defines an implicit conversion from Id string to PersonaId.
        /// </summary>
        /// <param name="uniqueId">The unique Id to convert to PersonaId.</param>
        /// <returns>A PersonaId initialized with the specified unique Id.</returns>
        public static implicit operator PersonaId(string uniqueId)
        {
            return new PersonaId(uniqueId);
        }

        /// <summary>
        /// Defines an implicit conversion from PersonaId to a Id string.
        /// </summary>
        /// <param name="PersonaId">The PersonaId to be converted</param>
        /// <returns>A PersonaId initialized with the specified unique Id.</returns>
        public static implicit operator String(PersonaId PersonaId)
        {
            if (PersonaId == null)
            {
                throw new ArgumentNullException("PersonaId");
            }

            if (String.IsNullOrEmpty(PersonaId.UniqueId))
            {
                return string.Empty;
            }
            else
            {
                // Ignoring the change key info
                return PersonaId.UniqueId;
            }
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.PersonaId;
        }

        /// <summary>
        /// Creates a new instance of PersonaId.
        /// </summary>
        /// <param name="uniqueId">The unique Id used to initialize the <see cref="PersonaId"/>.</param>
        public PersonaId(string uniqueId)
            : base(uniqueId)
        {
        }

        /// <summary>
        /// Gets a string representation of the Persona Id.
        /// </summary>
        /// <returns>The string representation of the Persona id.</returns>
        public override string ToString()
        {
            // We have ignored the change key portion
            return this.UniqueId;
        }
    }
}