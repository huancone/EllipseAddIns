package EllipseWebServicesClient;

import java.util.Collections;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.xml.namespace.QName;
import javax.xml.soap.*;
import javax.xml.ws.handler.MessageContext;
import javax.xml.ws.handler.soap.SOAPHandler;
import javax.xml.ws.handler.soap.SOAPMessageContext;

public class AuthenticationHandler implements SOAPHandler<SOAPMessageContext> {

    private final String namespaceURI = "http://schemas.xmlsoap.org/soap/envelope/";
    private final String uri = "http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd";

    @Override
    public Set<QName> getHeaders() {
        return Collections.EMPTY_SET;
    }

    @Override
    public boolean handleMessage(SOAPMessageContext context) {

        boolean Outbound;

        SOAPMessage msg = context.getMessage();
        Outbound = (Boolean) context.get(MessageContext.MESSAGE_OUTBOUND_PROPERTY);

        if (Outbound) {
            SOAPPart sp = msg.getSOAPPart();
            try {
                SOAPEnvelope env = sp.getEnvelope();
                env.addAttribute(new QName("Envelope"), namespaceURI);

                SOAPFactory soapFactory = SOAPFactory.newInstance();

                SOAPElement Security
                        = soapFactory.createElement(new QName(uri, "Security"));

                SOAPElement usernameToken
                        = soapFactory.createElement(new QName(uri, "UsernameToken"));

                SOAPElement Username
                        = soapFactory.createElement(new QName(uri, "Username"));
                Username.addTextNode(ClientConversation.username);

                SOAPElement Password
                        = soapFactory.createElement(new QName(uri, "Password"));
                Password.addTextNode(ClientConversation.password);

                usernameToken.addChildElement(Username);
                usernameToken.addChildElement(Password);
                Security.addChildElement(usernameToken);

                SOAPHeader soapHeader = env.addHeader();

                soapHeader.addChildElement(Security);
                //soapHeader.addAttribute(new QName(namespaceURI,"Header" ),"");
                return true;

            } catch (SOAPException ex) {
                Logger.getLogger(AuthenticationHandler.class.getName()).log(Level.SEVERE, null, ex);
                return false;
            }
        }
        return true;
    }

    @Override
    public boolean handleFault(SOAPMessageContext context) {
        return true;
    }

    @Override
    public void close(MessageContext context) {
    }

}
