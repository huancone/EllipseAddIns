using System;
using System.ServiceModel.Channels;
using System.ServiceModel.Dispatcher;
using System.ServiceModel;

namespace System.Web.Services.Ellipse
{
    public class EllipseMessageInspector : IDispatchMessageInspector, IClientMessageInspector
    {
        #region Message Inspector of the Service

        public object AfterReceiveRequest(ref Message request, IClientChannel channel, InstanceContext instanceContext)
        {
            return null;
        }

        public void BeforeSendReply(ref Message reply, object correlationState)
        {
        }

        #endregion

        #region Message Inspector of the Consumer

        public void AfterReceiveReply(ref Message reply, object correlationState)
        {
        }

        public object BeforeSendRequest(ref Message request, IClientChannel channel)
        {
            // Prepare the request message copy to be modified
            MessageBuffer buffer = request.CreateBufferedCopy(Int32.MaxValue);
            request = buffer.CreateMessage();

            // Simulate to have a random Key generation process
            request.Headers.Add(new SecurityHeader());
            request.Headers.Add(new ValueHeader("District", ClientConversation.district));
            request.Headers.Add(new ValueHeader("Position", ClientConversation.position));
            request.Headers.Add(new ValueHeader("Scope", ClientConversation.district));//TO DO New District Variable


            return null;
        }

        #endregion
    }
}
