// ---------------------------------------------------------------------------
// <copyright file="ManagementRole.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

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

        /// <summary>
        /// ToJsonObject
        /// </summary>
        /// <returns></returns>
        internal JsonObject ToJsonObject()
        {
            JsonObject managementRoleJsonObject = new JsonObject();
            if (this.userRoles != null)
            {
                managementRoleJsonObject.Add(XmlElementNames.UserRoles, this.userRoles);
            }

            if (this.applicationRoles != null)
            {
                managementRoleJsonObject.Add(XmlElementNames.ApplicationRoles, this.applicationRoles);
            }

            return managementRoleJsonObject;
        }
    }
}