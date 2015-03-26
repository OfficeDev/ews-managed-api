// ---------------------------------------------------------------------------
// <copyright file="GetUserUnifiedGroupsResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetUserUnifiedGroupsResponse class.</summary>
//-----------------------------------------------------------------------
namespace Microsoft.Exchange.WebServices.Data.Groups
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Text;

    /// <summary>
    /// Represents a response to a GetUserUnifiedGroupsResponse operation
    /// </summary>
    internal sealed class GetUserUnifiedGroupsResponse : ServiceResponse
    {
        /// <summary>
        /// The UnifiedGroups Sets associated with this response
        /// </summary>
        private Collection<UnifiedGroupsSet> groupsSets = new Collection<UnifiedGroupsSet>();

        /// <summary>
        /// Initializes a new instance of the <see cref="GetUserUnifiedGroupsResponse"/> class.
        /// </summary>
         internal GetUserUnifiedGroupsResponse() :
             base()
        {
        }

        /// <summary>
        /// Gets or sets the UnifiedGroupsSet associated with the response
        /// </summary>
         public Collection<UnifiedGroupsSet> GroupsSets
         { 
             get
             {
                 return this.groupsSets;
             }
         }

         /// <summary>
         /// Read Conversations from XML.
         /// </summary>
         /// <param name="reader">The reader.</param>
         internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
         {
             this.groupsSets.Clear();
             base.ReadElementsFromXml(reader);

             reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.GroupsSets);

             if (!reader.IsEmptyElement)
             {
                 reader.Read();
                 while (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.UnifiedGroupsSet))
                 {
                     UnifiedGroupsSet unifiedGroupsSet = new UnifiedGroupsSet();
                     unifiedGroupsSet.LoadFromXml(reader, XmlElementNames.UnifiedGroupsSet);
                     this.groupsSets.Add(unifiedGroupsSet);
                 }

                 // Skip end element GroupsSets
                 reader.EnsureCurrentNodeIsEndElement(XmlNamespace.NotSpecified, XmlElementNames.GroupsSets);
                 reader.Read();
             }
         }

         /// <summary>
         /// Reads response elements from Json.
         /// </summary>
         /// <param name="responseObject">The response object.</param>
         /// <param name="service">The service.</param>
         internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
         {
             this.groupsSets.Clear();
             base.ReadElementsFromJson(responseObject, service);

             if (responseObject.ContainsKey(XmlElementNames.GroupsSets))
             {
                 foreach (object unifiedGroupsSet in responseObject.ReadAsArray(XmlElementNames.UnifiedGroupsSet))
                 {
                     JsonObject jsonUnifiedGroupsSet = unifiedGroupsSet as JsonObject;
                     UnifiedGroupsSet unifiedGroupsSetResponse = new UnifiedGroupsSet();
                     unifiedGroupsSetResponse.LoadFromJson(jsonUnifiedGroupsSet, service);
                     this.groupsSets.Add(unifiedGroupsSetResponse);
                 }
             }
         }
    }
}