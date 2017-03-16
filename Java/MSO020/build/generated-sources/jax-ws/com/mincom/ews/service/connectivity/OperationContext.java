
package com.mincom.ews.service.connectivity;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;


/**
 * <p>Java class for OperationContext complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="OperationContext">
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *       &lt;sequence>
 *         &lt;element name="district" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *         &lt;element name="maxInstances" type="{http://www.w3.org/2001/XMLSchema}int" minOccurs="0"/>
 *         &lt;element name="position" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *         &lt;element name="returnWarnings" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="runAs" type="{http://connectivity.service.ews.mincom.com}RunAs" minOccurs="0"/>
 *         &lt;element name="trace" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="transaction" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "OperationContext", propOrder = {
    "district",
    "maxInstances",
    "position",
    "returnWarnings",
    "runAs",
    "trace",
    "transaction"
})
public class OperationContext {

    protected String district;
    protected Integer maxInstances;
    protected String position;
    protected Boolean returnWarnings;
    protected RunAs runAs;
    protected Boolean trace;
    protected String transaction;

    /**
     * Gets the value of the district property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getDistrict() {
        return district;
    }

    /**
     * Sets the value of the district property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setDistrict(String value) {
        this.district = value;
    }

    /**
     * Gets the value of the maxInstances property.
     * 
     * @return
     *     possible object is
     *     {@link Integer }
     *     
     */
    public Integer getMaxInstances() {
        return maxInstances;
    }

    /**
     * Sets the value of the maxInstances property.
     * 
     * @param value
     *     allowed object is
     *     {@link Integer }
     *     
     */
    public void setMaxInstances(Integer value) {
        this.maxInstances = value;
    }

    /**
     * Gets the value of the position property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPosition() {
        return position;
    }

    /**
     * Sets the value of the position property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPosition(String value) {
        this.position = value;
    }

    /**
     * Gets the value of the returnWarnings property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnWarnings() {
        return returnWarnings;
    }

    /**
     * Sets the value of the returnWarnings property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnWarnings(Boolean value) {
        this.returnWarnings = value;
    }

    /**
     * Gets the value of the runAs property.
     * 
     * @return
     *     possible object is
     *     {@link RunAs }
     *     
     */
    public RunAs getRunAs() {
        return runAs;
    }

    /**
     * Sets the value of the runAs property.
     * 
     * @param value
     *     allowed object is
     *     {@link RunAs }
     *     
     */
    public void setRunAs(RunAs value) {
        this.runAs = value;
    }

    /**
     * Gets the value of the trace property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isTrace() {
        return trace;
    }

    /**
     * Sets the value of the trace property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setTrace(Boolean value) {
        this.trace = value;
    }

    /**
     * Gets the value of the transaction property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getTransaction() {
        return transaction;
    }

    /**
     * Sets the value of the transaction property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setTransaction(String value) {
        this.transaction = value;
    }

}
