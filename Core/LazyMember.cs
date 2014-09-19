// ---------------------------------------------------------------------------
// <copyright file="LazyMember.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the LazyMember class.</summary>
//-----------------------------------------------------------------------

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
