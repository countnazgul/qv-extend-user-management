using System;
using System.Configuration;
using System.ServiceModel.Configuration;
using QV_UserManagement.ServiceSupport;

namespace QV_UserManagement.ServiceSupport
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