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
    /// Defines the index of a week day within a month.
    /// </summary>
    public enum DayOfTheWeekIndex
    {
        /// <summary>
        /// The first specific day of the week in the month. For example, the first Tuesday of the month. 
        /// </summary>
        First,

        /// <summary>
        /// The second specific day of the week in the month. For example, the second Tuesday of the month.
        /// </summary>
        Second,

        /// <summary>
        /// The third specific day of the week in the month. For example, the third Tuesday of the month.
        /// </summary>
        Third,

        /// <summary>
        /// The fourth specific day of the week in the month. For example, the fourth Tuesday of the month.
        /// </summary>
        Fourth,

        /// <summary>
        /// The last specific day of the week in the month. For example, the last Tuesday of the month.
        /// </summary>
        Last
    }
}