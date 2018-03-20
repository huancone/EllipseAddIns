
package com.mincom.enterpriseservice.screen;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.annotation.XmlElementDecl;
import javax.xml.bind.annotation.XmlRegistry;
import javax.xml.namespace.QName;
import com.mincom.enterpriseservice.exception.EnterpriseServiceException;
import com.mincom.enterpriseservice.exception.InvalidConnectionIdException;


/**
 * This object contains factory methods for each 
 * Java content interface and Java element interface 
 * generated in the com.mincom.enterpriseservice.screen package. 
 * <p>An ObjectFactory allows you to programatically 
 * construct new instances of the Java representation 
 * for XML content. The Java representation of XML 
 * content can consist of schema derived interfaces 
 * and classes representing the binding of schema 
 * type definitions, element declarations and model 
 * groups.  Factory methods for each of these are 
 * provided in this class.
 * 
 */
@XmlRegistry
public class ObjectFactory {

    private final static QName _EnterpriseServiceException_QNAME = new QName("http://screen.enterpriseservice.mincom.com", "EnterpriseServiceException");
    private final static QName _InvalidConnectionIdException_QNAME = new QName("http://screen.enterpriseservice.mincom.com", "InvalidConnectionIdException");

    /**
     * Create a new ObjectFactory that can be used to create new instances of schema derived classes for package: com.mincom.enterpriseservice.screen
     * 
     */
    public ObjectFactory() {
    }

    /**
     * Create an instance of {@link ExecuteScreen }
     * 
     */
    public ExecuteScreen createExecuteScreen() {
        return new ExecuteScreen();
    }

    /**
     * Create an instance of {@link PositionToMenu }
     * 
     */
    public PositionToMenu createPositionToMenu() {
        return new PositionToMenu();
    }

    /**
     * Create an instance of {@link Submit }
     * 
     */
    public Submit createSubmit() {
        return new Submit();
    }

    /**
     * Create an instance of {@link ScreenSubmitRequestDTO }
     * 
     */
    public ScreenSubmitRequestDTO createScreenSubmitRequestDTO() {
        return new ScreenSubmitRequestDTO();
    }

    /**
     * Create an instance of {@link PositionToMenuResponse }
     * 
     */
    public PositionToMenuResponse createPositionToMenuResponse() {
        return new PositionToMenuResponse();
    }

    /**
     * Create an instance of {@link SubmitResponse }
     * 
     */
    public SubmitResponse createSubmitResponse() {
        return new SubmitResponse();
    }

    /**
     * Create an instance of {@link ScreenDTO }
     * 
     */
    public ScreenDTO createScreenDTO() {
        return new ScreenDTO();
    }

    /**
     * Create an instance of {@link ExecuteScreenResponse }
     * 
     */
    public ExecuteScreenResponse createExecuteScreenResponse() {
        return new ExecuteScreenResponse();
    }

    /**
     * Create an instance of {@link ScreenNameValueDTO }
     * 
     */
    public ScreenNameValueDTO createScreenNameValueDTO() {
        return new ScreenNameValueDTO();
    }

    /**
     * Create an instance of {@link ArrayOfScreenFieldDTO }
     * 
     */
    public ArrayOfScreenFieldDTO createArrayOfScreenFieldDTO() {
        return new ArrayOfScreenFieldDTO();
    }

    /**
     * Create an instance of {@link ArrayOfScreenNameValueDTO }
     * 
     */
    public ArrayOfScreenNameValueDTO createArrayOfScreenNameValueDTO() {
        return new ArrayOfScreenNameValueDTO();
    }

    /**
     * Create an instance of {@link ScreenFieldDTO }
     * 
     */
    public ScreenFieldDTO createScreenFieldDTO() {
        return new ScreenFieldDTO();
    }

    /**
     * Create an instance of {@link JAXBElement }{@code <}{@link EnterpriseServiceException }{@code >}}
     * 
     */
    @XmlElementDecl(namespace = "http://screen.enterpriseservice.mincom.com", name = "EnterpriseServiceException")
    public JAXBElement<EnterpriseServiceException> createEnterpriseServiceException(EnterpriseServiceException value) {
        return new JAXBElement<EnterpriseServiceException>(_EnterpriseServiceException_QNAME, EnterpriseServiceException.class, null, value);
    }

    /**
     * Create an instance of {@link JAXBElement }{@code <}{@link InvalidConnectionIdException }{@code >}}
     * 
     */
    @XmlElementDecl(namespace = "http://screen.enterpriseservice.mincom.com", name = "InvalidConnectionIdException")
    public JAXBElement<InvalidConnectionIdException> createInvalidConnectionIdException(InvalidConnectionIdException value) {
        return new JAXBElement<InvalidConnectionIdException>(_InvalidConnectionIdException_QNAME, InvalidConnectionIdException.class, null, value);
    }

}
