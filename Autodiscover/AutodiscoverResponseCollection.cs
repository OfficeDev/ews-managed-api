// ---------------------------------------------------------------------------
// <copyright file="AutodiscoverResponseCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AutodiscoverResponseCollection class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;
    using System.Xml;
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Represents a collection of responses to a call to the Autodiscover service.
    /// </summary>
    /// <typeparam name="TResponse">The type of the responses in the collection.</typeparam>
    public abstract class AutodiscoverResponseCollection<TResponse> : AutodiscoverResponse, IEnumerable<TResponse>
        where TResponse : AutodiscoverResponse
    {
        private List<TResponse> responses;

        /// <summary>
        /// Initializes a new instance of the <see cref="AutodiscoverResponseCollection&lt;TResponse&gt;"/> class.
        /// </summary>
        internal AutodiscoverResponseCollection()
        {
            this.responses = new List<TResponse>();
        }
        
        /// <summary>
        /// Gets the number of responses in the collection.
        /// </summary>
        public int Count 
        {
            get { return this.responses.Count; }
        }

        /// <summary>
        /// Gets the response at the specified index.
        /// </summary>
        /// <param name="index">Index.</param>
        public TResponse this[int index]
        {
            get { return this.responses[index]; }
        }

        /// <summary>
        /// Gets the responses list.
        /// </summary>
        internal List<TResponse> Responses
        {
            get { return this.responses; }
        }

        /// <summary>
        /// Loads response from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="endElementName">End element name.</param>
        internal override void LoadFromXml(EwsXmlReader reader, string endElementName)
        {
            do
            {
                reader.Read();

                if (reader.NodeType == XmlNodeType.Element)
                {
                    if (reader.LocalName == this.GetResponseCollectionXmlElementName())
                    {
                        this.LoadResponseCollectionFromXml(reader);
                    }
                    else
                    {
                        base.LoadFromXml(reader, endElementName);
                    }
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Autodiscover, endElementName));
        }

        /// <summary>
        /// Loads the response collection from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        private void LoadResponseCollectionFromXml(EwsXmlReader reader)
        {
            if (!reader.IsEmptyElement)
            {
                do
                {
                    reader.Read();
                    if ((reader.NodeType == XmlNodeType.Element) && (reader.LocalName == this.GetResponseInstanceXmlElementName()))
                    {
                        TResponse response = this.CreateResponseInstance();
                        response.LoadFromXml(reader, this.GetResponseInstanceXmlElementName());
                        this.Responses.Add(response);
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Autodiscover, this.GetResponseCollectionXmlElementName()));
            }
        }

        /// <summary>
        /// Gets the name of the response collection XML element.
        /// </summary>
        /// <returns>Response collection XMl element name.</returns>
        internal abstract string GetResponseCollectionXmlElementName();

        /// <summary>
        /// Gets the name of the response instance XML element.
        /// </summary>
        /// <returns>Response instance XMl element name.</returns>
        internal abstract string GetResponseInstanceXmlElementName();

        /// <summary>
        /// Create a response instance.
        /// </summary>
        /// <returns>TResponse.</returns>
        internal abstract TResponse CreateResponseInstance();

        #region IEnumerable<TResponse>

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        public IEnumerator<TResponse> GetEnumerator()
        {
            return this.responses.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return (this.responses as System.Collections.IEnumerable).GetEnumerator();
        }

        #endregion
    }
}
