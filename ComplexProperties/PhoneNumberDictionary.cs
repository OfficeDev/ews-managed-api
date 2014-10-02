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
// <summary>Defines the PhoneNumberDictionary class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.ComponentModel;

    /// <summary>
    /// Represents a dictionary of phone numbers.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class PhoneNumberDictionary : DictionaryProperty<PhoneNumberKey, PhoneNumberEntry>
    {
        /// <summary>
        /// Gets the field URI.
        /// </summary>
        /// <returns>Field URI.</returns>
        internal override string GetFieldURI()
        {
            return "contacts:PhoneNumber";
        }

        /// <summary>
        /// Creates instance of dictionary entry.
        /// </summary>
        /// <returns>New instance.</returns>
        internal override PhoneNumberEntry CreateEntryInstance()
        {
            return new PhoneNumberEntry();
        }

        /// <summary>
        /// Gets or sets the phone number at the specified key.
        /// </summary>
        /// <param name="key">The key of the phone number to get or set.</param>
        /// <returns>The phone number at the specified key.</returns>
        public string this[PhoneNumberKey key]
        {
            get
            {
                return this.Entries[key].PhoneNumber;
            }

            set
            {
                if (value == null)
                {
                    this.InternalRemove(key);
                }
                else
                {
                    PhoneNumberEntry entry;

                    if (this.Entries.TryGetValue(key, out entry))
                    {
                        entry.PhoneNumber = value;
                        this.Changed();
                    }
                    else
                    {
                        entry = new PhoneNumberEntry(key, value);
                        this.InternalAdd(entry);
                    }
                }
            }
        }

        /// <summary>
        /// Tries to get the phone number associated with the specified key.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <param name="phoneNumber">
        /// When this method returns, contains the phone number associated with the specified key,
        /// if the key is found; otherwise, null. This parameter is passed uninitialized.
        /// </param>
        /// <returns>
        /// true if the Dictionary contains a phone number associated with the specified key; otherwise, false.
        /// </returns>
        public bool TryGetValue(PhoneNumberKey key, out string phoneNumber)
        {
            PhoneNumberEntry entry = null;

            if (this.Entries.TryGetValue(key, out entry))
            {
                phoneNumber = entry.PhoneNumber;

                return true;
            }
            else
            {
                phoneNumber = null;

                return false;
            }
        }
    }
}
