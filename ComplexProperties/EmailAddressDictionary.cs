// ---------------------------------------------------------------------------
// <copyright file="EmailAddressDictionary.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the EmailAddressDictionary class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.ComponentModel;

    /// <summary>
    /// Represents a dictionary of e-mail addresses.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class EmailAddressDictionary : DictionaryProperty<EmailAddressKey, EmailAddressEntry>
    {
        /// <summary>
        /// Gets the field URI.
        /// </summary>
        /// <returns>Field URI.</returns>
        internal override string GetFieldURI()
        {
            return "contacts:EmailAddress";
        }

        /// <summary>
        /// Creates instance of dictionary entry.
        /// </summary>
        /// <returns>New instance.</returns>
        internal override EmailAddressEntry CreateEntryInstance()
        {
            return new EmailAddressEntry();
        }

        /// <summary>
        /// Gets or sets the e-mail address at the specified key.
        /// </summary>
        /// <param name="key">The key of the e-mail address to get or set.</param>
        /// <returns>The e-mail address at the specified key.</returns>
        public EmailAddress this[EmailAddressKey key]
        {
            get
            {
                return this.Entries[key].EmailAddress;
            }

            set
            {
                if (value == null)
                {
                    this.InternalRemove(key);
                }
                else
                {
                    EmailAddressEntry entry;

                    if (this.Entries.TryGetValue(key, out entry))
                    {
                        entry.EmailAddress = value;
                        this.Changed();
                    }
                    else
                    {
                        entry = new EmailAddressEntry(key, value);
                        this.InternalAdd(entry);
                    }
                }
            }
        }

        /// <summary>
        /// Tries to get the e-mail address associated with the specified key.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <param name="emailAddress">
        /// When this method returns, contains the e-mail address associated with the specified key,
        /// if the key is found; otherwise, null. This parameter is passed uninitialized.
        /// </param>
        /// <returns>
        /// true if the Dictionary contains an e-mail address associated with the specified key; otherwise, false.
        /// </returns>
        public bool TryGetValue(EmailAddressKey key, out EmailAddress emailAddress)
        {
            EmailAddressEntry entry = null;

            if (this.Entries.TryGetValue(key, out entry))
            {
                emailAddress = entry.EmailAddress;

                return true;
            }
            else
            {
                emailAddress = null;

                return false;
            }
        }
    }
}