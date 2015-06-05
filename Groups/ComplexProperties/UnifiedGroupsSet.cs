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
    }
}