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
    }
}