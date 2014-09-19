// ---------------------------------------------------------------------------
// <copyright file="MobilePhone.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the MobilePhone class.</summary>
//-----------------------------------------------------------------------

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
