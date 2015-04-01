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
    /// Identifies the user configuration dictionary key and value types.
    /// </summary>
    public enum UserConfigurationDictionaryObjectType
    {
        /// <summary>
        /// DateTime type.
        /// </summary>
        DateTime,

        /// <summary>
        /// Boolean type.
        /// </summary>
        Boolean,

        /// <summary>
        /// Byte type.
        /// </summary>
        Byte,

        /// <summary>
        /// String type.
        /// </summary>
        String,

        /// <summary>
        /// 32-bit integer type.
        /// </summary>
        Integer32,

        /// <summary>
        /// 32-bit unsigned integer type.
        /// </summary>
        UnsignedInteger32,

        /// <summary>
        /// 64-bit integer type.
        /// </summary>
        Integer64,

        /// <summary>
        /// 64-bit unsigned integer type.
        /// </summary>
        UnsignedInteger64,

        /// <summary>
        /// String array type.
        /// </summary>
        StringArray,

        /// <summary>
        /// Byte array type
        /// </summary>
        ByteArray,
    }
}