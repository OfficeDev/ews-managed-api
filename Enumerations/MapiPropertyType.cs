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
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the MAPI type of an extended property.
    /// </summary>
    public enum MapiPropertyType
    {
        /// <summary>
        /// The property is of type ApplicationTime.
        /// </summary>
        ApplicationTime,

        /// <summary>
        /// The property is of type ApplicationTimeArray.
        /// </summary>
        ApplicationTimeArray,

        /// <summary>
        /// The property is of type Binary.
        /// </summary>
        Binary,

        /// <summary>
        /// The property is of type BinaryArray.
        /// </summary>
        BinaryArray,

        /// <summary>
        /// The property is of type Boolean.
        /// </summary>
        Boolean,

        /// <summary>
        /// The property is of type CLSID.
        /// </summary>
        CLSID,

        /// <summary>
        /// The property is of type CLSIDArray.
        /// </summary>
        CLSIDArray,

        /// <summary>
        /// The property is of type Currency.
        /// </summary>
        Currency,

        /// <summary>
        /// The property is of type CurrencyArray.
        /// </summary>
        CurrencyArray,

        /// <summary>
        /// The property is of type Double.
        /// </summary>
        Double,

        /// <summary>
        /// The property is of type DoubleArray.
        /// </summary>
        DoubleArray,

        /// <summary>
        /// The property is of type Error.
        /// </summary>
        Error,

        /// <summary>
        /// The property is of type Float.
        /// </summary>
        Float,

        /// <summary>
        /// The property is of type FloatArray.
        /// </summary>
        FloatArray,

        /// <summary>
        /// The property is of type Integer.
        /// </summary>
        Integer,

        /// <summary>
        /// The property is of type IntegerArray.
        /// </summary>
        IntegerArray,

        /// <summary>
        /// The property is of type Long.
        /// </summary>
        Long,

        /// <summary>
        /// The property is of type LongArray.
        /// </summary>
        LongArray,

        /// <summary>
        /// The property is of type Null.
        /// </summary>
        Null,

        /// <summary>
        /// The property is of type Object.
        /// </summary>
        Object,

        /// <summary>
        /// The property is of type ObjectArray.
        /// </summary>
        ObjectArray,

        /// <summary>
        /// The property is of type Short.
        /// </summary>
        Short,

        /// <summary>
        /// The property is of type ShortArray.
        /// </summary>
        ShortArray,

        /// <summary>
        /// The property is of type SystemTime.
        /// </summary>
        SystemTime,

        /// <summary>
        /// The property is of type SystemTimeArray.
        /// </summary>
        SystemTimeArray,

        /// <summary>
        /// The property is of type String.
        /// </summary>
        String,

        /// <summary>
        /// The property is of type StringArray.
        /// </summary>
        StringArray
    }
}