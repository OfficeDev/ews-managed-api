// ---------------------------------------------------------------------------
// <copyright file="PhysicalAddressDictionary.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

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
