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
// <summary>Defines the PhysicalAddressDictionary class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.ComponentModel;

    /// <summary>
    /// Represents a dictionary of physical addresses.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class PhysicalAddressDictionary : DictionaryProperty<PhysicalAddressKey, PhysicalAddressEntry>
    {
        /// <summary>
        /// Creates instance of dictionary entry.
        /// </summary>
        /// <returns>New instance.</returns>
        internal override PhysicalAddressEntry CreateEntryInstance()
        {
            return new PhysicalAddressEntry();
        }

        /// <summary>
        /// Gets or sets the physical address at the specified key.
        /// </summary>
        /// <param name="key">The key of the physical address to get or set.</param>
        /// <returns>The physical address at the specified key.</returns>
        public PhysicalAddressEntry this[PhysicalAddressKey key]
        {
            get
            {
                return this.Entries[key];
            }

            set
            {
                if (value == null)
                {
                    this.InternalRemove(key);
                }
                else
                {
                    value.Key = key;
                    this.InternalAddOrReplace(value);
                }
            }
        }

        /// <summary>
        /// Tries to get the physical address associated with the specified key.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <param name="physicalAddress">
        /// When this method returns, contains the physical address associated with the specified key,
        /// if the key is found; otherwise, null. This parameter is passed uninitialized.
        /// </param>
        /// <returns>
        /// true if the Dictionary contains a physical address associated with the specified key; otherwise, false.
        /// </returns>
        public bool TryGetValue(PhysicalAddressKey key, out PhysicalAddressEntry physicalAddress)
        {
            return this.Entries.TryGetValue(key, out physicalAddress);
        }
    }
}
