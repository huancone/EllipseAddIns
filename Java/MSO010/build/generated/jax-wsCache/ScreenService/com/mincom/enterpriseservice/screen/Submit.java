
package com.mincom.enterpriseservice.screen;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;
import javax.xml.bind.annotation.XmlType;
import com.mincom.ews.service.connectivity.OperationContext;


/**
 * <p>Java class for anonymous complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType>
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *       &lt;sequence>
 *         &lt;element name="context" type="{http://connectivity.service.ews.mincom.com}OperationContext"/>
 *         &lt;element name="screenSendRequestDTO" type="{http://screen.enterpriseservice.mincom.com}ScreenSubmitRequestDTO"/>
 *       &lt;/sequence>
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "", propOrder = {
    "context",
    "screenSendRequestDTO"
})
@XmlRootElement(name = "submit")
public class Submit {

    @XmlElement(required = true, nillable = true)
    protected OperationContext context;
    @XmlElement(required = true, nillable = true)
    protected ScreenSubmitRequestDTO screenSendRequestDTO;

    /**
     * Gets the value of the context property.
     * 
     * @return
     *     possible object is
     *     {@link OperationContext }
     *     
     */
    public OperationContext getContext() {
        return context;
    }

    /**
     * Sets the value of the context property.
     * 
     * @param value
     *     allowed object is
     *     {@link OperationContext }
     *     
     */
    public void setContext(OperationContext value) {
        this.context = value;
    }

    /**
     * Gets the value of the screenSendRequestDTO property.
     * 
     * @return
     *     possible object is
     *     {@link ScreenSubmitRequestDTO }
     *     
     */
    public ScreenSubmitRequestDTO getScreenSendRequestDTO() {
        return screenSendRequestDTO;
    }

    /**
     * Sets the value of the screenSendRequestDTO property.
     * 
     * @param value
     *     allowed object is
     *     {@link ScreenSubmitRequestDTO }
     *     
     */
    public void setScreenSendRequestDTO(ScreenSubmitRequestDTO value) {
        this.screenSendRequestDTO = value;
    }

}
