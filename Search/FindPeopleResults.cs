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

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Collections.ObjectModel;

    /// <summary>
    /// Represents the results of an Persona search operation.
    /// </summary>
    public sealed class FindPeopleResults
    {
        /// <summary>
        /// Creates a new instance of the <see cref="FindPeopleResults"/> class.
        /// </summary>
        internal FindPeopleResults()
        {
            this.Personas = new Collection<Persona>();
        }

        /// <summary>
        /// Accessors for the Personas that were found by the search operation.
        /// </summary>
        public Collection<Persona> Personas { get; internal set; }

        /// <summary>
        /// Accessors for the total count of Personas in view.
        /// </summary>
        public int? TotalCount { get; internal set; }

        /// <summary>
        /// Accessors for the first index of the matching row.
        /// </summary>
        public int? FirstMatchingRowIndex { get; internal set; }

        /// <summary>
        /// Accessors for the first index of the loaded row.
        /// </summary>
        public int? FirstLoadedRowIndex { get; internal set; }
    }
}