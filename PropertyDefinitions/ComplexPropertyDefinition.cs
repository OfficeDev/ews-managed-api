// ---------------------------------------------------------------------------
// <copyright file="ComplexPropertyDefinition.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ComplexPropertyDefinition class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Delegate used to create instances of ComplexProperty
    /// </summary>
    /// <typeparam name="TComplexProperty">Type of complex property.</typeparam>
    internal delegate TComplexProperty CreateComplexPropertyDelegate<TComplexProperty>()
        where TComplexProperty : ComplexProperty;

    /// <summary>
    /// Represents base complex property type.
    /// </summary>
    /// <typeparam name="TComplexProperty">The type of the complex property.</typeparam>
    internal class ComplexPropertyDefinition<TComplexProperty> : ComplexPropertyDefinitionBase
        where TComplexProperty : ComplexProperty
    {
        private CreateComplexPropertyDelegate<TComplexProperty> propertyCreationDelegate;

        /// <summary>
        /// Initializes a new instance of the <see cref="ComplexPropertyDefinition&lt;TComplexProperty&gt;"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="flags">The flags.</param>
        /// <param name="version">The version.</param>
        /// <param name="propertyCreationDelegate">Delegate used to create instances of ComplexProperty.</param>
        internal ComplexPropertyDefinition(
            string xmlElementName,
            PropertyDefinitionFlags flags,
            ExchangeVersion version,
            CreateComplexPropertyDelegate<TComplexProperty> propertyCreationDelegate)
            : base(
                xmlElementName,
                flags,
                version)
        {
            EwsUtilities.Assert(
                propertyCreationDelegate != null,
                "ComplexPropertyDefinition ctor",
                "CreateComplexPropertyDelegate cannot be null");

            this.propertyCreationDelegate = propertyCreationDelegate;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ComplexPropertyDefinition&lt;TComplexProperty&gt;"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="version">The version.</param>
        /// <param name="propertyCreationDelegate">Delegate used to create instances of ComplexProperty.</param>
        internal ComplexPropertyDefinition(
            string xmlElementName,
            string uri,
            ExchangeVersion version,
            CreateComplexPropertyDelegate<TComplexProperty> propertyCreationDelegate)
            : base(
                xmlElementName,
                uri,
                version)
        {
            this.propertyCreationDelegate = propertyCreationDelegate;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ComplexPropertyDefinition&lt;TComplexProperty&gt;"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="flags">The flags.</param>
        /// <param name="version">The version.</param>
        /// <param name="propertyCreationDelegate">Delegate used to create instances of ComplexProperty.</param>
        internal ComplexPropertyDefinition(
            string xmlElementName,
            string uri,
            PropertyDefinitionFlags flags,
            ExchangeVersion version,
            CreateComplexPropertyDelegate<TComplexProperty> propertyCreationDelegate)
            : base(
                xmlElementName,
                uri,
                flags,
                version)
        {
            this.propertyCreationDelegate = propertyCreationDelegate;
        }

        /// <summary>
        /// Creates the property instance.
        /// </summary>
        /// <param name="owner">The owner.</param>
        /// <returns>ComplexProperty instance.</returns>
        internal override ComplexProperty CreatePropertyInstance(ServiceObject owner)
        {
            TComplexProperty complexProperty = this.propertyCreationDelegate();
            IOwnedProperty ownedProperty = complexProperty as IOwnedProperty;

            if (ownedProperty != null)
            {
                ownedProperty.Owner = owner;
            }
            
            return complexProperty;
        }

        /// <summary>
        /// Gets the property type.
        /// </summary>
        public override Type Type
        {
            get { return typeof(TComplexProperty); }
        }
    }
}
