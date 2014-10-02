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
// ---------------------------------------------------------------------------
// <summary>
//      Versioning.cs
// </summary>
// ---------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Interface IDiscoveryVersionable
    /// This interface will be used to store versioning information on the request
    /// </summary>
    internal interface IDiscoveryVersionable
    {
        /// <summary>
        /// Gets or sets the server version.
        /// </summary>
        /// <value>
        /// The server version.
        /// </value>
        long ServerVersion { get; set; }
    }

    /// <summary>
    /// Class DiscoverySchemaChanges
    /// This class is a catalog of schema changes in discovery with the minimum server version in which they were introduced
    /// When making a schema change
    /// - First make the server side changes and check them in
    /// - Create SchemaChange() entry here for the change and the version at which it was checked int
    /// - In the request 
    ///     - Implement IDiscoveryVersionable
    ///     - In the Validate method verify if any new schema parameters are compatible if not error out
    ///     - In the WriteXml method downgrade the schema based on compatability checks
    /// Eg, SearchMailboxesRequest.cs
    /// </summary>
    internal static class DiscoverySchemaChanges
    {
        /// <summary>
        /// Initializes static members of the <see cref="DiscoverySchemaChanges"/> class.
        /// </summary>
        static DiscoverySchemaChanges()
        {
            // Schema change for passing extended data with the SearchMailboxes request
            SearchMailboxesExtendedData = new SchemaChange("15.0.730.0");

            // Schema change for additional search scopes such as "AllMailboxes", "PublicFolders", "SearchId" etc with the SearchMailboxes request
            SearchMailboxesAdditionalSearchScopes = new SchemaChange("15.0.730.0");
        }

        /// <summary>
        /// Gets the search mailboxes extended data.
        /// </summary>
        /// <value>
        /// The search mailboxes extended data.
        /// </value>
        internal static SchemaChange SearchMailboxesExtendedData { get; private set; }

        /// <summary>
        /// Gets the search mailboxes additional search scopes.
        /// </summary>
        /// <value>
        /// The search mailboxes additional search scopes.
        /// </value>
        internal static SchemaChange SearchMailboxesAdditionalSearchScopes { get; private set; }

        /// <summary>
        /// Class Feature
        /// </summary>
        internal sealed class SchemaChange
        {
            /// <summary>
            /// Gets the minimum server version.
            /// </summary>
            /// <value>
            /// The minimum server version.
            /// </value>
            internal long MinimumServerVersion { get; private set; }

            /// <summary>
            /// Initializes a new instance of the <see cref="SchemaChange"/> class.
            /// </summary>
            /// <param name="serverVersion">The server version.</param>
            internal SchemaChange(long serverVersion)
            {
                this.MinimumServerVersion = serverVersion;
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="SchemaChange"/> class.
            /// </summary>
            /// <param name="serverBuild">The server build.</param>
            internal SchemaChange(string serverBuild)
            {
                Version version = new Version(serverBuild);

                this.MinimumServerVersion = (version.Build & 0x7FFF) |
                                            ((version.Minor & 0x3F) << 16) |
                                            ((version.Major & 0x3F) << 22) |
                                            0x70008000;
            }

            /// <summary>
            /// Determines whether the specified versionable is compatible.
            /// </summary>
            /// <param name="versionable">The versionable.</param>
            /// <returns><c>true</c> if the specified versionable is compatible; otherwise, <c>false</c>.</returns>
            internal bool IsCompatible(IDiscoveryVersionable versionable)
            {
                // note: when ServerVersion is not set(i.e., => 0), we ignore compatible check on the client side. It will eventually fail server side schema check if incompatible
                return versionable.ServerVersion == 0 || versionable.ServerVersion >= MinimumServerVersion;
            }
        }
    }
}
