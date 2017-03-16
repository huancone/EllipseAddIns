
package com.mincom.enterpriseservice.ellipse.securityclass;

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
 *         &lt;element name="out" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}SecurityClassServiceModifyAttributesReplyCollectionDTO"/>
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
@XmlRootElement(name = "multipleModifyAttributesResponse")
public class MultipleModifyAttributesResponse {

    @XmlElement(required = true, nillable = true)
    protected SecurityClassServiceModifyAttributesReplyCollectionDTO out;

    /**
     * Gets the value of the out property.
     * 
     * @return
     *     possible object is
     *     {@link SecurityClassServiceModifyAttributesReplyCollectionDTO }
     *     
     */
    public SecurityClassServiceModifyAttributesReplyCollectionDTO getOut() {
        return out;
    }

    /**
     * Sets the value of the out property.
     * 
     * @param value
     *     allowed object is
     *     {@link SecurityClassServiceModifyAttributesReplyCollectionDTO }
     *     
     */
    public void setOut(SecurityClassServiceModifyAttributesReplyCollectionDTO value) {
        this.out = value;
    }

}
