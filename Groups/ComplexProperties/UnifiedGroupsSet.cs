// ---------------------------------------------------------------------------
// <copyright file="UnifiedGroupsSet.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the UnifiedGroupsSet class.</summary>
//-----------------------------------------------------------------------
namespace Microsoft.Exchange.WebServices.Data.Groups
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a UnifiedGroupsSet
    /// </summary>
    public class UnifiedGroupsSet : ComplexProperty
    {
        /// <summary>
        /// The list of unifiedGroups in this set.
        /// </summary>
        private List<UnifiedGroup> unifiedGroups = new List<UnifiedGroup>();

        /// <summary>
        /// Initializes a new instance of the <see cref="UnifiedGroupsSet"/> class.
        /// </summary>
        internal UnifiedGroupsSet() :
             base()
        {
        }

        /// <summary>
        /// Gets or sets the FilterType associated with this set
        /// </summary>
        public UnifiedGroupsFilterType FilterType { get; set; }

        /// <summary>
        /// Gets or sets the total groups in this set
        /// </summary>
        public int TotalGroups { get; set; }

        /// <summary>
        /// Gets the Groups contained in this set.
        /// </summary>
        public List<UnifiedGroup> Groups
        {
            get
            {
                return this.unifiedGroups;
            }
        }

         /// <summary>
         /// Read Conversations from XML.
         /// </summary>
         /// <param name="reader">The reader.</param>
         /// <param name="xmlElementName">The name of the xml element</param>
         internal override void LoadFromXml(EwsServiceXmlReader reader, string xmlElementName)
         {
             reader.EnsureCurrentNodeIsStartElement(XmlNamespace.Types, XmlElementNames.UnifiedGroupsSet);

             do
             {
                 reader.Read();
                 switch (reader.LocalName)
                 {
                     case XmlElementNames.FilterType:
                         this.FilterType = (UnifiedGroupsFilterType)Enum.Parse(typeof(UnifiedGroupsFilterType), reader.ReadElementValue(), false);
                         break;
                     case XmlElementNames.TotalGroups:
                         this.TotalGroups = reader.ReadElementValue<int>();
                         break;                     
                     case XmlElementNames.GroupsTag:
                         reader.Read();
                         while (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.UnifiedGroup))
                         {
                             UnifiedGroup unifiedGroup = new UnifiedGroup();
                             unifiedGroup.LoadFromXml(reader, XmlElementNames.UnifiedGroup);
                             this.unifiedGroups.Add(unifiedGroup);
                         }
                         
                         // Skip end element.
                         reader.EnsureCurrentNodeIsEndElement(XmlNamespace.NotSpecified, XmlElementNames.GroupsTag);
                         reader.Read();
                         break;
                     default:
                         break;
                 }
             }
             while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.UnifiedGroupsSet));

             // Skip end element
             reader.EnsureCurrentNodeIsEndElement(XmlNamespace.Types, XmlElementNames.UnifiedGroupsSet);
             reader.Read();
         }

         /// <summary>
         /// Reads response elements from Json.
         /// </summary>
         /// <param name="responseObject">The response object.</param>
         /// <param name="service">The service.</param>
         internal override void LoadFromJson(JsonObject responseObject, ExchangeService service)
         {
             if (responseObject.ContainsKey(XmlElementNames.FilterType))
             {
                 this.FilterType = (UnifiedGroupsFilterType)Enum.Parse(typeof(UnifiedGroupsFilterType), responseObject.ReadAsString(XmlElementNames.FilterType), false);
             }

             if (responseObject.ContainsKey(XmlElementNames.TotalGroups))
             {
                 this.TotalGroups = responseObject.ReadAsInt(XmlElementNames.TotalGroups);
             }

             if (responseObject.ContainsKey(XmlElementNames.GroupsTag))
             {
                 foreach (object unifiedGroup in responseObject.ReadAsArray(XmlElementNames.UnifiedGroup))
                 {
                     JsonObject jsonUnifiedGroup = unifiedGroup as JsonObject;
                     UnifiedGroup unifiedGroupResponse = new UnifiedGroup();
                     unifiedGroupResponse.LoadFromJson(jsonUnifiedGroup, service);
                     this.unifiedGroups.Add(unifiedGroupResponse);
                 }
             }
         }
    }
}