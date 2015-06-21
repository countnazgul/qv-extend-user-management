using System;
using System.Configuration;
using System.ServiceModel.Configuration;
using QlikviewEnhancedUserControl.ServiceSupport;

namespace QlikviewEnhancedUserControl.ServiceSupport
{
    public class ServiceKeyBehaviorExtensionElement : BehaviorExtensionElement
    {
        public override Type BehaviorType
        {
            get { return typeof(ServiceKeyEndpointBehavior); }
        }

        protected override object CreateBehavior()
        {
            return new ServiceKeyEndpointBehavior();
        }
    }
}