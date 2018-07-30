using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Microsoft.Exchange.WebServices.Data
{
    internal class UploadItemsRequest : MultiResponseServiceRequest<UploadItemsResponse>
    {
        IEnumerable<UploadItem> uploadItems;

        internal UploadItemsRequest(ExchangeService service, ServiceErrorHandling errorHandling) : base(service, errorHandling)
        {

        }

        internal override UploadItemsResponse CreateServiceResponse(ExchangeService service, int responseIndex)
        {
            return new UploadItemsResponse();
        }

        internal override string GetResponseMessageXmlElementName()
        {
            return XmlElementNames.UploadItemsResponseMessage;
        }

        internal override int GetExpectedResponseMessageCount()
        {
            return EwsUtilities.GetEnumeratedObjectCount(this.uploadItems);
        }

        internal override string GetXmlElementName()
        {
            return XmlElementNames.UploadItems;
        }

        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.UploadItemsResponse;
        }

        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2010;
        }

        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.Items);

            foreach(UploadItem uploadItem in uploadItems)
            {
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Item);
                writer.WriteAttributeString("CreateAction", uploadItem.CreateAction.ToString());

                uploadItem.ParentFolderId.WriteToXml(writer, XmlNamespace.Types, XmlElementNames.ParentFolderId);

                if(uploadItem.Id != null)
                {
                    uploadItem.Id.WriteToXml(writer);
                }

                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Data);
                writer.WriteBase64ElementValue(uploadItem.Data);
                writer.WriteEndElement();

                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }

        public IEnumerable<UploadItem> Items 
        {
            get { return uploadItems; }
            set { uploadItems = value; }
        }
    }
}
