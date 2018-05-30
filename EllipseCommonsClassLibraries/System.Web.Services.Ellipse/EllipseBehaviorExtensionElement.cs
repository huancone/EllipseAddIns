using System;
using System.ServiceModel.Configuration;

namespace System.Web.Services.Ellipse
{
    public class EllipseBehaviorExtensionElement : BehaviorExtensionElement
    {
        protected override object CreateBehavior()
        {
            return new EllipseHeaderBehavior();
        }

        public override Type BehaviorType
        {
            get { return typeof(EllipseHeaderBehavior); }
        }
    }
}
