#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

//-----------------------------------------------------------------------
// <summary>Defines the ServiceObject class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.ObjectModel;

    /// <summary>
    /// Represents the base abstract class for all item and folder types.
    /// </summary>
    public abstract class ServiceObject
    {
        private object lockObject = new object();
        private ExchangeService service;
        private PropertyBag propertyBag;
        private string xmlElementName;

        /// <summary>
        /// Triggers dispatch of the change event.
        /// </summary>
        internal void Changed()
        {
            if (this.OnChange != null)
            {
                this.OnChange(this);
            }
        }

        /// <summary>
        /// Throws exception if this is a new service object.
        /// </summary>
        internal void ThrowIfThisIsNew()
        {
            if (this.IsNew)
            {
                throw new InvalidOperationException(Strings.ServiceObjectDoesNotHaveId);
            }
        }

        /// <summary>
        /// Throws exception if this is not a new service object.
        /// </summary>
        internal void ThrowIfThisIsNotNew()
        {
            if (!this.IsNew)
            {
                throw new InvalidOperationException(Strings.ServiceObjectAlreadyHasId);
            }
        }

        /// <summary>
        /// This methods lets subclasses of ServiceObject override the default mechanism
        /// by which the XML element name associated with their type is retrieved.
        /// </summary>
        /// <returns>
        /// The XML element name associated with this type.
        /// If this method returns null or empty, the XML element name associated with this
        /// type is determined by the EwsObjectDefinition attribute that decorates the type,
        /// if present.
        /// </returns>
        /// <remarks>
        /// Item and folder classes that can be returned by EWS MUST rely on the EwsObjectDefinition
        /// attribute for XML element name determination.
        /// </remarks>
        internal virtual string GetXmlElementNameOverride()
        {
            return null;
        }

        /// <summary>
        /// GetXmlElementName retrieves the XmlElementName of this type based on the
        /// EwsObjectDefinition attribute that decorates it, if present.
        /// </summary>
        /// <returns>The XML element name associated with this type.</returns>
        internal string GetXmlElementName()
        {
            if (string.IsNullOrEmpty(this.xmlElementName))
            {
                this.xmlElementName = this.GetXmlElementNameOverride();

                if (string.IsNullOrEmpty(this.xmlElementName))
                {
                    lock (this.lockObject)
                    {
                        foreach (Attribute attribute in this.GetType().GetCustomAttributes(false))
                        {
                            ServiceObjectDefinitionAttribute definitionAttribute = attribute as ServiceObjectDefinitionAttribute;

                            if (definitionAttribute != null)
                            {
                                this.xmlElementName = definitionAttribute.XmlElementName;
                            }
                        }
                    }
                }
            }

            EwsUtilities.Assert(
                !string.IsNullOrEmpty(this.xmlElementName),
                "EwsObject.GetXmlElementName",
                string.Format("The class {0} does not have an associated XML element name.", this.GetType().Name));

            return this.xmlElementName;
        }

        /// <summary>
        /// Gets the name of the change XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal virtual string GetChangeXmlElementName()
        {
            return XmlElementNames.ItemChange;
        }

        /// <summary>
        /// Gets the name of the set field XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal virtual string GetSetFieldXmlElementName()
        {
            return XmlElementNames.SetItemField;
        }

        /// <summary>
        /// Gets the name of the delete field XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal virtual string GetDeleteFieldXmlElementName()
        {
            return XmlElementNames.DeleteItemField;
        }

        /// <summary>
        /// Gets a value indicating whether a time zone SOAP header should be emitted in a CreateItem
        /// or UpdateItem request so this item can be property saved or updated.
        /// </summary>
        /// <param name="isUpdateOperation">Indicates whether the operation being petrformed is an update operation.</param>
        /// <returns><c>true</c> if a time zone SOAP header should be emitted; otherwise, <c>false</c>.</returns>
        internal virtual bool GetIsTimeZoneHeaderRequired(bool isUpdateOperation)
        {
            return false;
        }

        /// <summary>
        /// Determines whether properties defined with ScopedDateTimePropertyDefinition require custom time zone scoping.
        /// </summary>
        /// <returns>
        ///     <c>true</c> if this item type requires custom scoping for scoped date/time properties; otherwise, <c>false</c>.
        /// </returns>
        internal virtual bool GetIsCustomDateTimeScopingRequired()
        {
            return false;
        }

        /// <summary>
        /// The property bag holding property values for this object.
        /// </summary>
        internal PropertyBag PropertyBag
        {
            get { return this.propertyBag; }
        }

        /// <summary>
        /// Internal constructor.
        /// </summary>
        /// <param name="service">EWS service to which this object belongs.</param>
        internal ServiceObject(ExchangeService service)
        {
            EwsUtilities.ValidateParam(service, "service");
            EwsUtilities.ValidateServiceObjectVersion(this, service.RequestedServerVersion);

            this.service = service;
            this.propertyBag = new PropertyBag(this);
        }

        /// <summary>
        /// Gets the schema associated with this type of object.
        /// </summary>
        public ServiceObjectSchema Schema
        {
            get { return this.GetSchema(); }
        }

        /// <summary>
        /// Internal method to return the schema associated with this type of object.
        /// </summary>
        /// <returns>The schema associated with this type of object.</returns>
        internal abstract ServiceObjectSchema GetSchema();

        /// <summary>
        /// Gets the minimum required server version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this service object type is supported.</returns>
        internal abstract ExchangeVersion GetMinimumRequiredServerVersion();

        /// <summary>
        /// Loads service object from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="clearPropertyBag">if set to <c>true</c> [clear property bag].</param>
        internal void LoadFromXml(EwsServiceXmlReader reader, bool clearPropertyBag)
        {
            this.PropertyBag.LoadFromXml(
                                reader,
                                clearPropertyBag,
                                null,       /* propertySet */
                                false);     /* summaryPropertiesOnly */
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal virtual void Validate()
        {
            this.PropertyBag.Validate();
        }

        /// <summary>
        /// Loads service object from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="clearPropertyBag">if set to <c>true</c> [clear property bag].</param>
        /// <param name="requestedPropertySet">The property set.</param>
        /// <param name="summaryPropertiesOnly">if set to <c>true</c> [summary props only].</param>
        internal void LoadFromXml(
                        EwsServiceXmlReader reader,
                        bool clearPropertyBag,
                        PropertySet requestedPropertySet,
                        bool summaryPropertiesOnly)
        {
            this.PropertyBag.LoadFromXml(
                                reader,
                                clearPropertyBag,
                                requestedPropertySet,
                                summaryPropertiesOnly);
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonServiceObject">The json service object.</param>
        /// <param name="service">The service.</param>
        /// <param name="clearPropertyBag">if set to <c>true</c> [clear property bag].</param>
        /// <param name="requestedPropertySet">The requested property set.</param>
        /// <param name="summaryPropertiesOnly">if set to <c>true</c> [summary properties only].</param>
        internal void LoadFromJson(JsonObject jsonServiceObject, ExchangeService service, bool clearPropertyBag, PropertySet requestedPropertySet, bool summaryPropertiesOnly)
        {
            this.PropertyBag.LoadFromJson(
                                jsonServiceObject,
                                service,
                                clearPropertyBag,
                                requestedPropertySet,
                                summaryPropertiesOnly);
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonObject">The json object.</param>
        /// <param name="service">The service.</param>
        /// <param name="clearPropertyBag">if set to <c>true</c> [clear property bag].</param>
        internal void LoadFromJson(JsonObject jsonObject, ExchangeService service, bool clearPropertyBag)
        {
            this.PropertyBag.LoadFromJson(
                                jsonObject,
                                service,
                                clearPropertyBag,
                                null,       /* propertySet */
                                false);     /* summaryPropertiesOnly */
        }

        /// <summary>
        /// Clears the object's change log.
        /// </summary>
        internal void ClearChangeLog()
        {
            this.PropertyBag.ClearChangeLog();
        }

        /// <summary>
        /// Writes service object as XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal void WriteToXml(EwsServiceXmlWriter writer)
        {
            this.PropertyBag.WriteToXml(writer);
        }

        /// <summary>
        /// Creates a JSON representation of this object.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="isUpdateOperation">if set to <c>true</c> [is update operation].</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        internal object ToJson(ExchangeService service, bool isUpdateOperation)
        {
            return this.PropertyBag.ToJson(service, isUpdateOperation);
        }

        /// <summary>
        /// Writes service object for update as XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal void WriteToXmlForUpdate(EwsServiceXmlWriter writer)
        {
            this.PropertyBag.WriteToXmlForUpdate(writer);
        }

        /// <summary>
        /// Writes service object for update as Json.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns></returns>
        internal object WriteToJsonForUpdate(ExchangeService service)
        {
            return this.PropertyBag.ToJson(service, true);
        }

        /// <summary>
        /// Loads the specified set of properties on the object.
        /// </summary>
        /// <param name="propertySet">The properties to load.</param>
        internal abstract void InternalLoad(PropertySet propertySet);

        /// <summary>
        /// Deletes the object.
        /// </summary>
        /// <param name="deleteMode">The deletion mode.</param>
        /// <param name="sendCancellationsMode">Indicates whether meeting cancellation messages should be sent.</param>
        /// <param name="affectedTaskOccurrences">Indicate which occurrence of a recurring task should be deleted.</param>
        internal abstract void InternalDelete(
            DeleteMode deleteMode,
            SendCancellationsMode? sendCancellationsMode,
            AffectedTaskOccurrence? affectedTaskOccurrences);

        /// <summary>
        /// Loads the specified set of properties. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="propertySet">The properties to load.</param>
        public void Load(PropertySet propertySet)
        {
            this.InternalLoad(propertySet);
        }

        /// <summary>
        /// Loads the first class properties. Calling this method results in a call to EWS.
        /// </summary>
        public void Load()
        {
            this.InternalLoad(PropertySet.FirstClassProperties);
        }

        /// <summary>
        /// Gets the value of specified property in this instance. 
        /// </summary>
        /// <param name="propertyDefinition">Definition of the property to get.</param>
        /// <exception cref="ServiceVersionException">Raised if this property requires a later version of Exchange.</exception>
        /// <exception cref="PropertyException">Raised if this property hasn't been assigned or loaded. Raised for set if property cannot be updated or deleted.</exception>
        public object this[PropertyDefinitionBase propertyDefinition]
        {
            get
            {
                object propertyValue;

                PropertyDefinition propDef = propertyDefinition as PropertyDefinition;
                if (propDef != null)
                {
                    return this.PropertyBag[propDef];
                }
                else
                {
                    ExtendedPropertyDefinition extendedPropDef = propertyDefinition as ExtendedPropertyDefinition;
                    if (extendedPropDef != null)
                    {
                        if (this.TryGetExtendedProperty(extendedPropDef, out propertyValue))
                        {
                            return propertyValue;
                        }
                        else
                        {
                            throw new ServiceObjectPropertyException(Strings.MustLoadOrAssignPropertyBeforeAccess, propertyDefinition);
                        }
                    }
                    else
                    {
                        // Other subclasses of PropertyDefinitionBase are not supported.
                        throw new NotSupportedException(string.Format(
                            Strings.OperationNotSupportedForPropertyDefinitionType,
                            propertyDefinition.GetType().Name));
                    }
                }
            }
        }

        /// <summary>
        /// Try to get the value of a specified extended property in this instance. 
        /// </summary>
        /// <param name="propertyDefinition">The property definition.</param>
        /// <param name="propertyValue">The property value.</param>
        /// <typeparam name="T">Type of expected property value.</typeparam>
        /// <returns>True if property retrieved, false otherwise.</returns>
        internal bool TryGetExtendedProperty<T>(ExtendedPropertyDefinition propertyDefinition, out T propertyValue)
        {
            ExtendedPropertyCollection propertyCollection = this.GetExtendedProperties();

            if ((propertyCollection != null) &&
                propertyCollection.TryGetValue<T>(propertyDefinition, out propertyValue))
            {
                return true;
            }
            else
            {
                propertyValue = default(T);
                return false;
            }
        }

        /// <summary>
        /// Try to get the value of a specified property in this instance. 
        /// </summary>
        /// <param name="propertyDefinition">The property definition.</param>
        /// <param name="propertyValue">The property value.</param>
        /// <returns>True if property retrieved, false otherwise.</returns>
        public bool TryGetProperty(PropertyDefinitionBase propertyDefinition, out object propertyValue)
        {
            return this.TryGetProperty<object>(propertyDefinition, out propertyValue);
        }

        /// <summary>
        /// Try to get the value of a specified property in this instance. 
        /// </summary>
        /// <param name="propertyDefinition">The property definition.</param>
        /// <param name="propertyValue">The property value.</param>
        /// <typeparam name="T">Type of expected property value.</typeparam>
        /// <returns>True if property retrieved, false otherwise.</returns>
        public bool TryGetProperty<T>(PropertyDefinitionBase propertyDefinition, out T propertyValue)
        {
            PropertyDefinition propDef = propertyDefinition as PropertyDefinition;
            if (propDef != null)
            {
                return this.PropertyBag.TryGetProperty<T>(propDef, out propertyValue);
            }
            else
            {
                ExtendedPropertyDefinition extPropDef = propertyDefinition as ExtendedPropertyDefinition;
                if (extPropDef != null)
                {
                    return this.TryGetExtendedProperty<T>(extPropDef, out propertyValue);
                }
                else
                {
                    // Other subclasses of PropertyDefinitionBase are not supported.
                    throw new NotSupportedException(string.Format(
                        Strings.OperationNotSupportedForPropertyDefinitionType,
                        propertyDefinition.GetType().Name));
                }
            }
        }

        /// <summary>
        /// Gets the collection of loaded property definitions.
        /// </summary>
        /// <returns>Collection of property definitions.</returns>
        public Collection<PropertyDefinitionBase> GetLoadedPropertyDefinitions()
        {
            Collection<PropertyDefinitionBase> propDefs = new Collection<PropertyDefinitionBase>();
            foreach (PropertyDefinition propDef in this.PropertyBag.Properties.Keys)
            {
                propDefs.Add(propDef);
            }

            if (this.GetExtendedProperties() != null)
            {
                foreach (ExtendedProperty extProp in this.GetExtendedProperties())
                {
                    propDefs.Add(extProp.PropertyDefinition);
                }
            }

            return propDefs;
        }

        /// <summary>
        /// Gets the ExchangeService the object is bound to.
        /// </summary>
        public ExchangeService Service
        {
            get { return this.service; }
            internal set { this.service = value; }
        }

        /// <summary>
        /// The property definition for the Id of this object.
        /// </summary>
        /// <returns>A PropertyDefinition instance.</returns>
        internal virtual PropertyDefinition GetIdPropertyDefinition()
        {
            return null;
        }

        /// <summary>
        /// The unique Id of this object.
        /// </summary>
        /// <returns>A ServiceId instance.</returns>
        internal ServiceId GetId()
        {
            PropertyDefinition idPropertyDefinition = this.GetIdPropertyDefinition();

            object serviceId = null;

            if (idPropertyDefinition != null)
            {
                this.PropertyBag.TryGetValue(idPropertyDefinition, out serviceId);
            }

            return (ServiceId)serviceId;
        }

        /// <summary>
        /// Indicates whether this object is a real store item, or if it's a local object
        /// that has yet to be saved.
        /// </summary>
        public virtual bool IsNew
        {
            get
            {
                ServiceId id = this.GetId();

                return id == null ? true : !id.IsValid;
            }
        }

        /// <summary>
        /// Gets a value indicating whether the object has been modified and should be saved.
        /// </summary>
        public bool IsDirty
        {
            get { return this.PropertyBag.IsDirty; }
        }

        /// <summary>
        /// Gets the extended properties collection.
        /// </summary>
        /// <returns>Extended properties collection.</returns>
        internal virtual ExtendedPropertyCollection GetExtendedProperties()
        {
            return null;
        }

        /// <summary>
        /// Defines an event that is triggered when the service object changes.
        /// </summary>
        internal event ServiceObjectChangedDelegate OnChange;
    }
}