// ---------------------------------------------------------------------------
// <copyright file="ComparisonHelpers.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ComparisonHelpers class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    using System;
    using System.Collections;

    /// <summary>
    /// Represents a set of helper methods for performing string comparisons.
    /// </summary>
    internal static class ComparisonHelpers
    {
        /// <summary>
        /// Case insensitive check if the collection contains the string.
        /// </summary>
        /// <param name="collection">The collection of objects, only strings are checked</param>
        /// <param name="match">String to match</param>
        /// <returns>true, if match contained in the collection</returns>
        internal static bool CaseInsensitiveContains(this ICollection collection, string match)
        {
            foreach (object obj in collection)
            {
                string str = obj as string;
                if (str != null)
                {
                    if (string.Compare(str, match, StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        return true;
                    }
                }
            }

            return false;
        }
    }
}
