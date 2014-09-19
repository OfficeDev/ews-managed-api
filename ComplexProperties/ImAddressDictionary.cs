// ---------------------------------------------------------------------------
// <copyright file="ImAddressDictionary.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ImAddressDictionary class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.ComponentModel;

    /// <summary>
    /// Represents a dictionary of Instant Messaging addresses.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class ImAddressDictionary : DictionaryProperty<ImAddressKey, ImAddressEntry>
    {
        /// <summary>
        /// Gets the field URI.
        /// </summary>
        /// <returns>Field URI.</returns>
        internal override string GetFieldURI()
        {
            return "contacts:ImAddress";
        }

        /// <summary>
        /// Creates instance of dictionary entry.
        /// </summary>
        /// <returns>New instance.</returns>
        internal override ImAddressEntry CreateEntryInstance()
        {
            return new ImAddressEntry();
        }

        /// <summary>
        /// Gets or sets the Instant Messaging address at the specified key.
        /// </summary>
        /// <param name="key">The key of the Instant Messaging address to get or set.</param>
        /// <returns>The Instant Messaging address at the specified key.</returns>
        public string this[ImAddressKey key]
        {
            get
            {
                return this.Entries[key].ImAddress;
            }

            set
            {
                if (value == null)
                {
                    this.InternalRemove(key);
                }
                else
                {
                    ImAddressEntry entry;

                    if (this.Entries.TryGetValue(key, out entry))
                    {
                        entry.ImAddress = value;
                        this.Changed();
                    }
                    else
                    {
                        entry = new ImAddressEntry(key, value);
                        this.InternalAdd(entry);
                    }
                }
            }
        }

        /// <summary>
        /// Tries to get the IM address associated with the specified key.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <param name="imAddress">
        /// When this method returns, contains the IM address associated with the specified key,
        /// if the key is found; otherwise, null. This parameter is passed uninitialized.
        /// </param>
        /// <returns>
        /// true if the Dictionary contains an IM address associated with the specified key; otherwise, false.
        /// </returns>
        public bool TryGetValue(ImAddressKey key, out string imAddress)
        {
            ImAddressEntry entry = null;

            if (this.Entries.TryGetValue(key, out entry))
            {
                imAddress = entry.ImAddress;

                return true;
            }
            else
            {
                imAddress = null;

                return false;
            }
        }
    }
}
