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
// <summary>Defines the PhoneCall class.</summary>
//-----------------------------------------------------------------------
namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Text;

    /// <summary>
    /// Represents a phone call.
    /// </summary>
    public sealed class PhoneCall : ComplexProperty
    {
        private const string SuccessfulResponseText = "OK";
        private const int SuccessfulResponseCode = 200;

        private ExchangeService service;
        private PhoneCallState state;
        private ConnectionFailureCause connectionFailureCause;
        private string sipResponseText;
        private int sipResponseCode;
        private PhoneCallId id;
        
        /// <summary>
        /// PhoneCall Constructor.
        /// </summary>
        /// <param name="service">EWS service to which this object belongs.</param>
        internal PhoneCall(ExchangeService service)
        {
            EwsUtilities.Assert(
                service != null,
                "PhoneCall.ctor",
                "service is null");

            this.service = service;
            this.state = PhoneCallState.Connecting;
            this.connectionFailureCause = ConnectionFailureCause.None;
            this.sipResponseText = PhoneCall.SuccessfulResponseText;
            this.sipResponseCode = PhoneCall.SuccessfulResponseCode;          
        }

        /// <summary>
        /// PhoneCall Constructor.
        /// </summary>
        /// <param name="service">EWS service to which this object belongs.</param>
        /// <param name="id">The Id of the phone call.</param>
        internal PhoneCall(ExchangeService service, PhoneCallId id)
            : this(service)
        {
            this.id = id;    
        }

        /// <summary>
        /// Refreshes the state of this phone call.
        /// </summary>
        public void Refresh()
        {
            PhoneCall phoneCall = service.UnifiedMessaging.GetPhoneCallInformation(this.id);
            this.state = phoneCall.State;
            this.connectionFailureCause = phoneCall.ConnectionFailureCause;
            this.sipResponseText = phoneCall.SIPResponseText;
            this.sipResponseCode = phoneCall.SIPResponseCode;
        }

        /// <summary>
        /// Disconnects this phone call.
        /// </summary>
        public void Disconnect()
        {
            // If call is already disconnected, throw exception
            //
            if (this.state == PhoneCallState.Disconnected)
            {
                throw new ServiceLocalException(Strings.PhoneCallAlreadyDisconnected);
            }

            this.service.UnifiedMessaging.DisconnectPhoneCall(this.id);
            this.state = PhoneCallState.Disconnected;
        }

        /// <summary>
        /// Tries to read an element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.PhoneCallState:
                    this.state = reader.ReadElementValue<PhoneCallState>();
                    return true;
                case XmlElementNames.ConnectionFailureCause:
                    this.connectionFailureCause = reader.ReadElementValue<ConnectionFailureCause>();
                    return true;
                case XmlElementNames.SIPResponseText:
                    this.sipResponseText = reader.ReadElementValue();
                    return true;
                case XmlElementNames.SIPResponseCode:
                    this.sipResponseCode = reader.ReadElementValue<int>();
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service"></param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            foreach (string key in jsonProperty.Keys)
            {
                switch (key)
                {
                    case XmlElementNames.PhoneCallState:
                        this.state = jsonProperty.ReadEnumValue<PhoneCallState>(key);
                        break;
                    case XmlElementNames.ConnectionFailureCause:
                        this.connectionFailureCause = jsonProperty.ReadEnumValue<ConnectionFailureCause>(key);
                        break;
                    case XmlElementNames.SIPResponseText:
                        this.sipResponseText = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.SIPResponseCode:
                        this.sipResponseCode = jsonProperty.ReadAsInt(key);
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Gets a value indicating the last known state of this phone call.
        /// </summary>
        public PhoneCallState State
        {
            get
            {
                return this.state;
            }
        }

        /// <summary>
        /// Gets a value indicating the reason why this phone call failed to connect.
        /// </summary>
        public ConnectionFailureCause ConnectionFailureCause
        {
            get
            {
                return this.connectionFailureCause;
            }
        }

        /// <summary>
        /// Gets the SIP response text of this phone call.
        /// </summary>
        public string SIPResponseText
        {
            get
            {
                return this.sipResponseText;
            }
        }

        /// <summary>
        /// Gets the SIP response code of this phone call.
        /// </summary>
        public int SIPResponseCode
        {
            get
            {
                return this.sipResponseCode;
            }
        }
    }
}
