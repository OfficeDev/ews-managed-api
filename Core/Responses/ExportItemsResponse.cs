using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Microsoft.Exchange.WebServices.Data
{
    public class ExportItemsResponse : ServiceResponse
    {
        private ItemId itemId = new ItemId();
        private byte[] data;

        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);

            if(!reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.ItemId))
            {
                reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.ItemId);
            }

            itemId.LoadFromXml(reader, XmlNamespace.Messages, XmlAttributeNames.ItemId);

            if(!reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.Data))
            {
                reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.Data);
            }
            data = reader.ReadBase64ElementValue();
        }

        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            base.ReadElementsFromJson(responseObject, service);
        }

        public ItemId ItemId
        {
            get
            {
                return itemId;
            }
        }

        public byte[] Data
        {
            get
            {
                return data;
            }
        }

    }
}
