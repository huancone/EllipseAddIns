
package com.mincom.enterpriseservice.exception;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.ArrayOfErrorMessageDTO;


/**
 * <p>Java class for EnterpriseServiceOperationException complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="EnterpriseServiceOperationException">
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *       &lt;sequence>
 *         &lt;element name="errorMessages" type="{http://ellipse.enterpriseservice.mincom.com}ArrayOfErrorMessageDTO" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "EnterpriseServiceOperationException", propOrder = {
    "errorMessages"
})
public class EnterpriseServiceOperationException {

    protected ArrayOfErrorMessageDTO errorMessages;

    /**
     * Gets the value of the errorMessages property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfErrorMessageDTO }
     *     
     */
    public ArrayOfErrorMessageDTO getErrorMessages() {
        return errorMessages;
    }

    /**
     * Sets the value of the errorMessages property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfErrorMessageDTO }
     *     
     */
    public void setErrorMessages(ArrayOfErrorMessageDTO value) {
        this.errorMessages = value;
    }

}
