// ---------------------------------------------------------------------------
// <copyright file="EwsServiceXmlReader.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the EwsServiceXmlReader class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// XML reader.
    /// </summary>
    internal class EwsServiceXmlReader : EwsXmlReader
    {
        #region Private members

        private ExchangeService service;

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="EwsServiceXmlReader"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="service">The service.</param>
        internal EwsServiceXmlReader(Stream stream, ExchangeService service)
            : base(stream)
        {
            this.service = service;
        }

        #endregion

        /// <summary>
        /// Converts the specified string into a DateTime objects.
        /// </summary>
        /// <param name="dateTimeString">The date time string to convert.</param>
        /// <returns>A DateTime representing the converted string.</returns>
        private DateTime? ConvertStringToDateTime(string dateTimeString)
        {
            return this.Service.ConvertUniversalDateTimeStringToLocalDateTime(dateTimeString);
        }

        /// <summary>
        /// Converts the specified string into a unspecified Date object, ignoring offset.
        /// </summary>
        /// <param name="dateTimeString">The date time string to convert.</param>
        /// <returns>A DateTime representing the converted string.</returns>
        private DateTime? ConvertStringToUnspecifiedDate(string dateTimeString)
        {
            return this.Service.ConvertStartDateToUnspecifiedDateTime(dateTimeString);
        }

        /// <summary>
        /// Reads the element value as date time.
        /// </summary>
        /// <returns>Element value.</returns>
        public DateTime? ReadElementValueAsDateTime()
        {
            return this.ConvertStringToDateTime(this.ReadElementValue());
        }

        /// <summary>
        /// Reads the element value as unspecified date.
        /// </summary>
        /// <returns>Element value.</returns>
        public DateTime? ReadElementValueAsUnspecifiedDate()
        {
            return this.ConvertStringToUnspecifiedDate(this.ReadElementValue());
        }

        /// <summary>
        /// Reads the element value as date time, assuming it is unbiased (e.g. 2009/01/01T08:00) 
        /// and scoped to service's time zone.
        /// </summary>
        /// <returns>The element's value as a DateTime object.</returns>
        public DateTime ReadElementValueAsUnbiasedDateTimeScopedToServiceTimeZone()
        {
            string elementValue = this.ReadElementValue();
            return EwsUtilities.ParseAsUnbiasedDatetimescopedToServicetimeZone(elementValue, this.Service);
        }

        /// <summary>
        /// Reads the element value as date time.
        /// </summary>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <param name="localName">Name of the local.</param>
        /// <returns>Element value.</returns>
        public DateTime? ReadElementValueAsDateTime(XmlNamespace xmlNamespace, string localName)
        {
            return this.ConvertStringToDateTime(this.ReadElementValue(xmlNamespace, localName));
        }

        /// <summary>
        /// Reads the service objects collection from XML.
        /// </summary>
        /// <typeparam name="TServiceObject">The type of the service object.</typeparam>
        /// <param name="collectionXmlNamespace">Namespace of the collection XML element.</param>
        /// <param name="collectionXmlElementName">Name of the collection XML element.</param>
        /// <param name="getObjectInstanceDelegate">The get object instance delegate.</param>
        /// <param name="clearPropertyBag">if set to <c>true</c> [clear property bag].</param>
        /// <param name="requestedPropertySet">The requested property set.</param>
        /// <param name="summaryPropertiesOnly">if set to <c>true</c> [summary properties only].</param>
        /// <returns>List of service objects.</returns>
        public List<TServiceObject> ReadServiceObjectsCollectionFromXml<TServiceObject>(
            XmlNamespace collectionXmlNamespace,
            string collectionXmlElementName,
            GetObjectInstanceDelegate<TServiceObject> getObjectInstanceDelegate,
            bool clearPropertyBag,
            PropertySet requestedPropertySet,
            bool summaryPropertiesOnly) where TServiceObject : ServiceObject
        {
            List<TServiceObject> serviceObjects = new List<TServiceObject>();
            TServiceObject serviceObject = null;

            if (!this.IsStartElement(collectionXmlNamespace, collectionXmlElementName))
            {
                this.ReadStartElement(collectionXmlNamespace, collectionXmlElementName);
            }

            if (!this.IsEmptyElement)
            {
                do
                {
                    this.Read();

                    if (this.IsStartElement())
                    {
                        serviceObject = getObjectInstanceDelegate(this.Service, this.LocalName);

                        if (serviceObject == null)
                        {
                            this.SkipCurrentElement();
                        }
                        else
                        {
                            if (string.Compare(this.LocalName, serviceObject.GetXmlElementName(), StringComparison.Ordinal) != 0)
                            {
                                throw new ServiceLocalException(
                                    string.Format(
                                        "The type of the object in the store ({0}) does not match that of the local object ({1}).",
                                        this.LocalName,
                                        serviceObject.GetXmlElementName()));
                            }

                            serviceObject.LoadFromXml(
                                            this,
                                            clearPropertyBag,
                                            requestedPropertySet,
                                            summaryPropertiesOnly);

                            serviceObjects.Add(serviceObject);
                        }
                    }
                }
                while (!this.IsEndElement(collectionXmlNamespace, collectionXmlElementName));
            }

            return serviceObjects;
        }

        /// <summary>
        /// Reads the service objects collection from XML.
        /// </summary>
        /// <typeparam name="TServiceObject">The type of the service object.</typeparam>
        /// <param name="collectionXmlElementName">Name of the collection XML element.</param>
        /// <param name="getObjectInstanceDelegate">The get object instance delegate.</param>
        /// <param name="clearPropertyBag">if set to <c>true</c> [clear property bag].</param>
        /// <param name="requestedPropertySet">The requested property set.</param>
        /// <param name="summaryPropertiesOnly">if set to <c>true</c> [summary properties only].</param>
        /// <returns>List of service objects.</returns>
        public List<TServiceObject> ReadServiceObjectsCollectionFromXml<TServiceObject>(
            string collectionXmlElementName,
            GetObjectInstanceDelegate<TServiceObject> getObjectInstanceDelegate,
            bool clearPropertyBag,
            PropertySet requestedPropertySet,
            bool summaryPropertiesOnly) where TServiceObject : ServiceObject
        {
            return this.ReadServiceObjectsCollectionFromXml<TServiceObject>(
                                XmlNamespace.Messages,
                                collectionXmlElementName,
                                getObjectInstanceDelegate,
                                clearPropertyBag,
                                requestedPropertySet,
                                summaryPropertiesOnly);
        }

        /// <summary>
        /// Gets the service.
        /// </summary>
        /// <value>The service.</value>
        public ExchangeService Service
        {
            get { return this.service; }
        }
    }
}