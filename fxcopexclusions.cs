// ---------------------------------------------------------------------------
// <copyright file="FxCopExclusions.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>
//  FxCop exclusions
// </summary>
//-----------------------------------------------------------------------

using System.Diagnostics.CodeAnalysis;

[module: SuppressMessage("Exchange.Usage", "EX0031:DoNotUseUnsafeXmlParsers", Scope = "type", Target = "Microsoft.Exchange.WebServices.Data.EwsServiceMultiResponseXmlReader")]

[module: SuppressMessage("Exchange.Usage", "EX0031:DoNotUseUnsafeXmlParsers", Scope = "type", Target = "Microsoft.Exchange.WebServices.Data.EwsXmlReader")]
