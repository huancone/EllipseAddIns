
package com.mincom.enterpriseservice.ellipse.employee;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;
import javax.xml.bind.annotation.XmlType;


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
 *         &lt;element name="out" type="{http://employee.ellipse.enterpriseservice.mincom.com}EmployeeServiceTransferPositionReplyDTO"/>
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
    "out"
})
@XmlRootElement(name = "transferPositionResponse")
public class TransferPositionResponse {

    @XmlElement(required = true, nillable = true)
    protected EmployeeServiceTransferPositionReplyDTO out;

    /**
     * Gets the value of the out property.
     * 
     * @return
     *     possible object is
     *     {@link EmployeeServiceTransferPositionReplyDTO }
     *     
     */
    public EmployeeServiceTransferPositionReplyDTO getOut() {
        return out;
    }

    /**
     * Sets the value of the out property.
     * 
     * @param value
     *     allowed object is
     *     {@link EmployeeServiceTransferPositionReplyDTO }
     *     
     */
    public void setOut(EmployeeServiceTransferPositionReplyDTO value) {
        this.out = value;
    }

}
