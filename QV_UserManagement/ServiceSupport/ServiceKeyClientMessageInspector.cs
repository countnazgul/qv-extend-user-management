using System.ServiceModel;
using System.ServiceModel.Channels;
using System.ServiceModel.Dispatcher;

namespace QlikviewEnhancedUserControl.ServiceSupport
{

    public class ServiceKeyClientMessageInspector : IClientMessageInspector
    {
        private const string SERVICE_KEY_HTTP_HEADER = "X-Service-Key";

        public static string ServiceKey { get; set; }

        public object BeforeSendRequest(ref Message request, IClientChannel channel)
        {
            object httpRequestMessageObject;
            if (request.Properties.TryGetValue(HttpRequestMessageProperty.Name, out httpRequestMessageObject))
            {
                HttpRequestMessageProperty httpRequestMessage = httpRequestMessageObject as HttpRequestMessageProperty;
                if (httpRequestMessage != null)
                {
                    httpRequestMessage.Headers[SERVICE_KEY_HTTP_HEADER] = (ServiceKey ?? string.Empty);
                }
                else
                {
                    httpRequestMessage = new HttpRequestMessageProperty();
                    httpRequestMessage.Headers.Add(SERVICE_KEY_HTTP_HEADER, (ServiceKey ?? string.Empty));
                    request.Properties[HttpRequestMessageProperty.Name] = httpRequestMessage;
                }
            }
            else
            {
                HttpRequestMessageProperty httpRequestMessage = new HttpRequestMessageProperty();
                httpRequestMessage.Headers.Add(SERVICE_KEY_HTTP_HEADER, (ServiceKey ?? string.Empty));
                request.Properties.Add(HttpRequestMessageProperty.Name, httpRequestMessage);
            }
            return null;
        }

        public void AfterReceiveReply(ref Message reply, object correlationState) { }
    }
}