
package com.mincom.enterpriseservice.ellipse.securityclass;

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
 *         &lt;element name="requestParameters" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}ArrayOfSecurityClassServiceReadRequestDTO"/>
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
    "requestParameters"
})
@XmlRootElement(name = "multipleReadWithPartialResults")
public class MultipleReadWithPartialResults {

    @XmlElement(required = true, nillable = true)
    protected OperationContext context;
    @XmlElement(required = true, nillable = true)
    protected ArrayOfSecurityClassServiceReadRequestDTO requestParameters;

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
     * Gets the value of the requestParameters property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfSecurityClassServiceReadRequestDTO }
     *     
     */
    public ArrayOfSecurityClassServiceReadRequestDTO getRequestParameters() {
        return requestParameters;
    }

    /**
     * Sets the value of the requestParameters property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfSecurityClassServiceReadRequestDTO }
     *     
     */
    public void setRequestParameters(ArrayOfSecurityClassServiceReadRequestDTO value) {
        this.requestParameters = value;
    }

}
