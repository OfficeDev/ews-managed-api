// ---------------------------------------------------------------------------
// <copyright file="SearchFolderSchema.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SearchFolderSchema class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Diagnostics.CodeAnalysis;

    /// <summary>
    /// Represents the schema for search folders.
    /// </summary>
    [Schema]
    public class SearchFolderSchema : FolderSchema
    {
        /// <summary>
        /// Field URIs for search folders.
        /// </summary>
        private static class FieldUris
        {
            public const string SearchParameters = "folder:SearchParameters";
        }

        /// <summary>
        /// Defines the SearchParameters property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition SearchParameters =
            new ComplexPropertyDefinition<SearchFolderParameters>(
                XmlElementNames.SearchParameters,
                FieldUris.SearchParameters,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.AutoInstantiateOnRead,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new SearchFolderParameters(); });

        // This must be declared after the property definitions
        internal static new readonly SearchFolderSchema Instance = new SearchFolderSchema();

        /// <summary>
        /// Registers properties.
        /// </summary>
        /// <remarks>
        /// IMPORTANT NOTE: PROPERTIES MUST BE REGISTERED IN SCHEMA ORDER (i.e. the same order as they are defined in types.xsd)
        /// </remarks>
        internal override void RegisterProperties()
        {
            base.RegisterProperties();

            this.RegisterProperty(SearchParameters);
        }
    }
}
