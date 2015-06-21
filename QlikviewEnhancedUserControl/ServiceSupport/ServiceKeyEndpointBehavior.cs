using System.ServiceModel.Channels;
using System.ServiceModel.Description;
using System.ServiceModel.Dispatcher;

namespace QlikviewEnhancedUserControl.ServiceSupport
{
    public class ServiceKeyEndpointBehavior : IEndpointBehavior
    {
        public void Validate(ServiceEndpoint endpoint) { }

        public void AddBindingParameters(ServiceEndpoint endpoint, BindingParameterCollection bindingParameters) { }

        public void ApplyDispatchBehavior(ServiceEndpoint endpoint, EndpointDispatcher endpointDispatcher) { }

        public void ApplyClientBehavior(ServiceEndpoint endpoint, ClientRuntime clientRuntime)
        {
            clientRuntime.MessageInspectors.Add(new ServiceKeyClientMessageInspector());
        }
    }
}