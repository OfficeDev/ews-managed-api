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
    /// Delegate called to perform the actual initialization of the member
    /// </summary>
    /// <typeparam name="T">Wrapped lazy member type</typeparam>
    /// <returns>Newly instantiated and initialized member</returns>
    internal delegate T InitializeLazyMember<T>();

    /// <summary>
    /// Wrapper class for lazy members.  Does lazy initialization of member on first access.
    /// </summary>
    /// <typeparam name="T">Type of the lazy member</typeparam>
    /// <remarks>If we find ourselves creating a whole bunch of these in our code, we need to rethink
    /// this.  Each lazy member holds the actual member, a lock object, a boolean flag and a delegate.
    /// That can turn into a whole lot of overhead.</remarks>
    internal class LazyMember<T>
    {
        private T lazyMember;
        private InitializeLazyMember<T> initializationDelegate;
        private object lockObject = new object();
        private bool initialized = false;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="initializationDelegate">The initialization delegate to call for the item on first access
        /// </param>
        public LazyMember(InitializeLazyMember<T> initializationDelegate)
        {
            this.initializationDelegate = initializationDelegate;
        }

        /// <summary>
        /// Public accessor for the lazy member.  Lazy initializes the member on first access
        /// </summary>
        public T Member
        {
            get
            {
                if (!this.initialized)
                {
                    lock (this.lockObject)
                    {
                        if (!this.initialized)
                        {
                            this.lazyMember = this.initializationDelegate();
                        }
                        this.initialized = true;
                    }
                }
                return this.lazyMember;
            }
        }
    }
}