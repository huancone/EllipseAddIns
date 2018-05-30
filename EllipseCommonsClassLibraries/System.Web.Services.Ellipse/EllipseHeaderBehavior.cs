using System;
using System.ServiceModel.Dispatcher;
using System.ServiceModel.Description;

namespace System.Web.Services.Ellipse
{
    [AttributeUsage(AttributeTargets.Class)]
    public class EllipseHeaderBehavior : Attribute, IEndpointBehavior
    {
        #region IEndpointBehavior Members

        public void AddBindingParameters(ServiceEndpoint endpoint, System.ServiceModel.Channels.BindingParameterCollection bindingParameters)
        {
        }

        public void ApplyClientBehavior(ServiceEndpoint endpoint, System.ServiceModel.Dispatcher.ClientRuntime clientRuntime)
        {
            clientRuntime.MessageInspectors.Add(new EllipseMessageInspector());
        }

        public void ApplyDispatchBehavior(ServiceEndpoint endpoint, System.ServiceModel.Dispatcher.EndpointDispatcher endpointDispatcher)
        {
            ChannelDispatcher channelDispatcher = endpointDispatcher.ChannelDispatcher;
            if (channelDispatcher != null)
            {
                foreach (EndpointDispatcher ed in channelDispatcher.Endpoints)
                {
                    ed.DispatchRuntime.MessageInspectors.Add(new EllipseMessageInspector());
                }
            }
        }

        public void Validate(ServiceEndpoint endpoint)
        {
        }

        #endregion
    }
}
