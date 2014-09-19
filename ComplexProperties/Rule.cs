// ---------------------------------------------------------------------------
// <copyright file="Rule.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the Rule class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents a rule that automatically handles incoming messages.
    /// A rule consists of a set of conditions and exceptions that determine whether or 
    /// not a set of actions should be executed on incoming messages.
    /// </summary>
    public sealed class Rule : ComplexProperty
    {
        /// <summary>
        /// The rule ID.
        /// </summary>
        private string ruleId;

        /// <summary>
        /// The rule display name.
        /// </summary>
        private string displayName;

        /// <summary>
        /// The rule priority.
        /// </summary>
        private int priority;

        /// <summary>
        /// The rule status of enabled or not.
        /// </summary>
        private bool isEnabled;

        /// <summary>
        /// The rule status of is supported or not.
        /// </summary>
        private bool isNotSupported;

        /// <summary>
        /// The rule status of in error or not.
        /// </summary>
        private bool isInError;
        
        /// <summary>
        /// The rule conditions. 
        /// </summary>
        private RulePredicates conditions;

        /// <summary>
        /// The rule actions. 
        /// </summary>
        private RuleActions actions;
        
        /// <summary>
        /// The rule exceptions. 
        /// </summary>
        private RulePredicates exceptions;

        /// <summary>
        /// Initializes a new instance of the <see cref="Rule"/> class.
        /// </summary>
        public Rule()
            : base()
        {
            //// New rule has priority as 0 by default
            this.priority = 1;
            //// New rule is enabled by default
            this.isEnabled = true;
            this.conditions = new RulePredicates();
            this.actions = new RuleActions();
            this.exceptions = new RulePredicates();
        }

        /// <summary>
        /// Gets or sets the Id of this rule.
        /// </summary>
        public string Id
        {
            get
            {
                return this.ruleId;
            }

            set
            {
                this.SetFieldValue<string>(ref this.ruleId, value);
            }
        }

        /// <summary>
        /// Gets or sets the name of this rule as it should be displayed to the user.
        /// </summary>
        public string DisplayName
        {
            get
            {
                return this.displayName;
            }

            set
            {
                this.SetFieldValue<string>(ref this.displayName, value);
            }
        }

        /// <summary>
        /// Gets or sets the priority of this rule, which determines its execution order.
        /// </summary>
        public int Priority
        {
            get
            {
                return this.priority;
            }

            set
            {
                this.SetFieldValue<int>(ref this.priority, value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether this rule is enabled.
        /// </summary>
        public bool IsEnabled
        {
            get
            {
                return this.isEnabled;
            }

            set
            {
                this.SetFieldValue<bool>(ref this.isEnabled, value);
            }
        }

        /// <summary>
        /// Gets a value indicating whether this rule can be modified via EWS. 
        /// If IsNotSupported is true, the rule cannot be modified via EWS.
        /// </summary>
        public bool IsNotSupported
        {
            get
            {
                return this.isNotSupported;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether this rule has errors. A rule that is in error 
        /// cannot be processed unless it is updated and the error is corrected.
        /// </summary>
        public bool IsInError
        {
            get
            {
                return this.isInError;
            }

            set
            {
                this.SetFieldValue<bool>(ref this.isInError, value);
            }
        }

        /// <summary>
        /// Gets the conditions that determine whether or not this rule should be
        /// executed against incoming messages.
        /// </summary>
        public RulePredicates Conditions
        {
            get
            {
                return this.conditions;
            }
        }

        /// <summary>
        /// Gets the actions that should be executed against incoming messages if the
        /// conditions evaluate as true.
        /// </summary>
        public RuleActions Actions
        {
            get
            {
                return this.actions;
            }
        }

        /// <summary>
        /// Gets the exceptions that determine if this rule should be skipped even if 
        /// its conditions evaluate to true.
        /// </summary>
        public RulePredicates Exceptions
        {
            get
            {
                return this.exceptions;
            }
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.DisplayName:
                    this.displayName = reader.ReadElementValue();
                    return true;
                case XmlElementNames.RuleId:
                    this.ruleId = reader.ReadElementValue();
                    return true;
                case XmlElementNames.Priority:
                    this.priority = reader.ReadElementValue<int>();
                    return true;
                case XmlElementNames.IsEnabled:
                    this.isEnabled = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsNotSupported:
                    this.isNotSupported = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsInError:
                    this.isInError = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.Conditions:
                    this.conditions.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.Actions:
                    this.actions.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.Exceptions:
                    this.exceptions.LoadFromXml(reader, reader.LocalName);
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service">The service.</param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            foreach (string key in jsonProperty.Keys)
            {
                switch (key)
                {
                    case XmlElementNames.DisplayName:
                        this.displayName = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.RuleId:
                        this.ruleId = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.Priority:
                        this.priority = jsonProperty.ReadAsInt(key);
                        break;
                    case XmlElementNames.IsEnabled:
                        this.isEnabled = jsonProperty.ReadAsBool(key);
                        break;
                    case XmlElementNames.IsNotSupported:
                        this.isNotSupported = jsonProperty.ReadAsBool(key);
                        break;
                    case XmlElementNames.IsInError:
                        this.isInError = jsonProperty.ReadAsBool(key);
                        break;
                    case XmlElementNames.Conditions:
                        this.conditions.LoadFromJson(jsonProperty, service);
                        break;
                    case XmlElementNames.Actions:
                        this.actions.LoadFromJson(jsonProperty, service);
                        break;
                    case XmlElementNames.Exceptions:
                        this.exceptions.LoadFromJson(jsonProperty, service);
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            if (!string.IsNullOrEmpty(this.Id))
            {
                writer.WriteElementValue(
                    XmlNamespace.Types, 
                    XmlElementNames.RuleId, 
                    this.Id);
            }

            writer.WriteElementValue(
                XmlNamespace.Types, 
                XmlElementNames.DisplayName, 
                this.DisplayName);
            writer.WriteElementValue(
                XmlNamespace.Types, 
                XmlElementNames.Priority, 
                this.Priority);
            writer.WriteElementValue(
                XmlNamespace.Types, 
                XmlElementNames.IsEnabled, 
                this.IsEnabled);
            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.IsInError,
                this.IsInError);
            this.Conditions.WriteToXml(writer, XmlElementNames.Conditions);
            this.Exceptions.WriteToXml(writer, XmlElementNames.Exceptions);
            this.Actions.WriteToXml(writer, XmlElementNames.Actions);
        }

        /// <summary>
        /// Serializes the property to a Json value.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        internal override object InternalToJson(ExchangeService service)
        {
            JsonObject jsonProperty = new JsonObject();

            if (!string.IsNullOrEmpty(this.Id))
            {
                jsonProperty.Add(XmlElementNames.RuleId, this.Id);
            }

            jsonProperty.Add(XmlElementNames.DisplayName, this.DisplayName);
            jsonProperty.Add(XmlElementNames.Priority, this.Priority);
            jsonProperty.Add(XmlElementNames.IsEnabled, this.IsEnabled);
            jsonProperty.Add(XmlElementNames.IsInError, this.IsInError);

            jsonProperty.Add(XmlElementNames.Conditions, this.Conditions.InternalToJson(service));
            jsonProperty.Add(XmlElementNames.Exceptions, this.Exceptions.InternalToJson(service));
            jsonProperty.Add(XmlElementNames.Actions, this.Actions.InternalToJson(service));

            return jsonProperty;
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal override void InternalValidate()
        {
            base.InternalValidate();
            EwsUtilities.ValidateParam(this.displayName, "DisplayName");
            EwsUtilities.ValidateParam(this.conditions, "Conditions");
            EwsUtilities.ValidateParam(this.exceptions, "Exceptions");
            EwsUtilities.ValidateParam(this.actions, "Actions");
        }
    }
}
