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
    /// <summary>
    /// Represents a mobile phone.
    /// </summary>
    public sealed class MobilePhone : ISelfValidate
    {
        /// <summary>
        /// Name of the mobile phone.
        /// </summary>
        private string name;

        /// <summary>
        /// Phone number of the mobile phone.
        /// </summary>
        private string phoneNumber;

        /// <summary>
        /// Initializes a new instance of the <see cref="MobilePhone"/> class.
        /// </summary>
        public MobilePhone()
        {
        }
        
        /// <summary>
        /// Initializes a new instance of the <see cref="MobilePhone"/> class.
        /// </summary>
        /// <param name="name">The name associated with the mobile phone.</param>
        /// <param name="phoneNumber">The mobile phone number.</param>
        public MobilePhone(string name, string phoneNumber)
        {
            this.name = name;
            this.phoneNumber = phoneNumber;
        }

        /// <summary>
        /// Gets or sets the name associated with this mobile phone.
        /// </summary>
        public string Name
        {
            get { return this.name; }
            set { this.name = value; }
        }

        /// <summary>
        /// Gets or sets the number of this mobile phone.
        /// </summary>
        public string PhoneNumber
        {
            get { return this.phoneNumber; }
            set { this.phoneNumber = value; }
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        void ISelfValidate.Validate()
        {
            if (string.IsNullOrEmpty(this.PhoneNumber))
            {
                throw new ServiceValidationException("PhoneNumber cannot be empty.");
            }
        }
    }
}