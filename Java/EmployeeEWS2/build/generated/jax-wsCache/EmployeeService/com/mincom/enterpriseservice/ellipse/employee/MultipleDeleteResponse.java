
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
 *         &lt;element name="out" type="{http://employee.ellipse.enterpriseservice.mincom.com}EmployeeServiceDeleteReplyCollectionDTO"/>
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
@XmlRootElement(name = "multipleDeleteResponse")
public class MultipleDeleteResponse {

    @XmlElement(required = true, nillable = true)
    protected EmployeeServiceDeleteReplyCollectionDTO out;

    /**
     * Gets the value of the out property.
     * 
     * @return
     *     possible object is
     *     {@link EmployeeServiceDeleteReplyCollectionDTO }
     *     
     */
    public EmployeeServiceDeleteReplyCollectionDTO getOut() {
        return out;
    }

    /**
     * Sets the value of the out property.
     * 
     * @param value
     *     allowed object is
     *     {@link EmployeeServiceDeleteReplyCollectionDTO }
     *     
     */
    public void setOut(EmployeeServiceDeleteReplyCollectionDTO value) {
        this.out = value;
    }

}
