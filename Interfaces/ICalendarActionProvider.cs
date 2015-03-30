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
    /// Interface defintion of a group of methods that are common to items that return CalendarActionResults
    /// </summary>
    internal interface ICalendarActionProvider
    {
        /// <summary>
        /// Implements the Accept method.
        /// </summary>
        /// <param name="sendResponse">Indicates whether to send a response to the organizer.</param>
        /// <returns>A CalendarActionResults object containing the various items that were created or modified as a result of this operation.</returns>
        CalendarActionResults Accept(bool sendResponse);

        /// <summary>
        /// Implements the AcceptTentatively method.
        /// </summary>
        /// <param name="sendResponse">Indicates whether to send a response to the organizer.</param>
        /// <returns>A CalendarActionResults object containing the various items that were created or modified as a result of this operation.</returns>
        CalendarActionResults AcceptTentatively(bool sendResponse);

        /// <summary>
        /// Implements the Decline method.
        /// </summary>
        /// <param name="sendResponse">Indicates whether to send a response to the organizer.</param>
        /// <returns>A CalendarActionResults object containing the various items that were created or modified as a result of this operation.</returns>
        CalendarActionResults Decline(bool sendResponse);

        /// <summary>
        /// Implements the CreateAcceptMessage method.
        /// </summary>
        /// <param name="tentative">Indicates whether the new AcceptMeetingInvitationMessage should represent a Tentative accept response (as opposed to an Accept response).</param>
        /// <returns>A new AcceptMeetingInvitationMessage.</returns>
        AcceptMeetingInvitationMessage CreateAcceptMessage(bool tentative);

        /// <summary>
        /// Implements the DeclineMeetingInvitationMessage method.
        /// </summary>
        /// <returns>A new DeclineMeetingInvitationMessage.</returns>
        DeclineMeetingInvitationMessage CreateDeclineMessage();
    }
}