
package com.mincom.enterpriseservice.ellipse.employee;

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
 *         &lt;element name="requestParameters" type="{http://employee.ellipse.enterpriseservice.mincom.com}EmployeeServiceRetrieveForExtractRequestDTO"/>
 *         &lt;element name="requiredAttributes" type="{http://employee.ellipse.enterpriseservice.mincom.com}EmployeeServiceRetrieveForExtractRequiredAttributesDTO"/>
 *         &lt;element name="restartInfo" type="{http://www.w3.org/2001/XMLSchema}string"/>
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
    "requestParameters",
    "requiredAttributes",
    "restartInfo"
})
@XmlRootElement(name = "retrieveForExtract")
public class RetrieveForExtract {

    @XmlElement(required = true, nillable = true)
    protected OperationContext context;
    @XmlElement(required = true, nillable = true)
    protected EmployeeServiceRetrieveForExtractRequestDTO requestParameters;
    @XmlElement(required = true, nillable = true)
    protected EmployeeServiceRetrieveForExtractRequiredAttributesDTO requiredAttributes;
    @XmlElement(required = true, nillable = true)
    protected String restartInfo;

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
     *     {@link EmployeeServiceRetrieveForExtractRequestDTO }
     *     
     */
    public EmployeeServiceRetrieveForExtractRequestDTO getRequestParameters() {
        return requestParameters;
    }

    /**
     * Sets the value of the requestParameters property.
     * 
     * @param value
     *     allowed object is
     *     {@link EmployeeServiceRetrieveForExtractRequestDTO }
     *     
     */
    public void setRequestParameters(EmployeeServiceRetrieveForExtractRequestDTO value) {
        this.requestParameters = value;
    }

    /**
     * Gets the value of the requiredAttributes property.
     * 
     * @return
     *     possible object is
     *     {@link EmployeeServiceRetrieveForExtractRequiredAttributesDTO }
     *     
     */
    public EmployeeServiceRetrieveForExtractRequiredAttributesDTO getRequiredAttributes() {
        return requiredAttributes;
    }

    /**
     * Sets the value of the requiredAttributes property.
     * 
     * @param value
     *     allowed object is
     *     {@link EmployeeServiceRetrieveForExtractRequiredAttributesDTO }
     *     
     */
    public void setRequiredAttributes(EmployeeServiceRetrieveForExtractRequiredAttributesDTO value) {
        this.requiredAttributes = value;
    }

    /**
     * Gets the value of the restartInfo property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getRestartInfo() {
        return restartInfo;
    }

    /**
     * Sets the value of the restartInfo property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setRestartInfo(String value) {
        this.restartInfo = value;
    }

}
