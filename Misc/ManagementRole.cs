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
    using System.Linq;

    /// <summary>
    /// ManagementRoles
    /// </summary>
    public sealed class ManagementRoles
    {
        private readonly string[] userRoles;
        private readonly string[] applicationRoles;

        /// <summary>
        /// Initializes a new instance of the <see cref="ManagementRoles"/> class.
        /// </summary>
        /// <param name="userRole"></param>
        public ManagementRoles(string userRole)
        {
            EwsUtilities.ValidateParam(userRole, "userRole");
            this.userRoles = new string[] { userRole };
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ManagementRoles"/> class.
        /// </summary>
        /// <param name="userRole"></param>
        /// <param name="applicationRole"></param>
        public ManagementRoles(string userRole, string applicationRole)
        {
            if (userRole != null)
            {
                EwsUtilities.ValidateParam(userRole, "userRole");
                this.userRoles = new string[] { userRole };
            }

            if (applicationRole != null)
            {
                EwsUtilities.ValidateParam(applicationRole, "applicationRole");
                this.applicationRoles = new string[] { applicationRole };
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ManagementRoles"/> class.
        /// </summary>
        /// <param name="userRoles"></param>
        /// <param name="applicationRoles"></param>
        public ManagementRoles(string[] userRoles, string[] applicationRoles)
        {
            if (userRoles != null)
            {
                this.userRoles = userRoles.ToArray();
            }

            if (applicationRoles != null)
            {
                this.applicationRoles = applicationRoles.ToArray();
            }
        }

        /// <summary>
        /// WriteToXml
        /// </summary>
        /// <param name="writer"></param>
        internal void WriteToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.ManagementRole);
            WriteRolesToXml(writer, this.userRoles, XmlElementNames.UserRoles);
            WriteRolesToXml(writer, this.applicationRoles, XmlElementNames.ApplicationRoles);
            writer.WriteEndElement();
        }

        /// <summary>
        /// WriteRolesToXml
        /// </summary>
        /// <param name="writer"></param>
        /// <param name="roles"></param>
        /// <param name="elementName"></param>
        private void WriteRolesToXml(EwsServiceXmlWriter writer, string[] roles, string elementName)
        {
            if (roles != null)
            {
                writer.WriteStartElement(XmlNamespace.Types, elementName);

                foreach (string role in roles)
                {
                    writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Role, role);
                }

                writer.WriteEndElement();
            }
        }
    }
}