
package com.mincom.enterpriseservice.ellipse;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlSeeAlso;
import javax.xml.bind.annotation.XmlType;
import com.mincom.ellipse.attribute.ArrayOfAttribute;
import com.mincom.enterpriseservice.ellipse.securityclass.SecurityClassServiceModifyAttributesRequestDTO;
import com.mincom.enterpriseservice.ellipse.securityclass.SecurityClassServiceModifyMethodsRequestDTO;
import com.mincom.enterpriseservice.ellipse.securityclass.SecurityClassServiceModifyRequestDTO;
import com.mincom.enterpriseservice.ellipse.securityclass.SecurityClassServiceReadRequestDTO;
import com.mincom.enterpriseservice.ellipse.securityclass.SecurityClassServiceRetrieveClassesRequestDTO;
import com.mincom.enterpriseservice.ellipse.securityclass.SecurityClassServiceRetrieveRequestDTO;


/**
 * <p>Java class for AbstractDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="AbstractDTO">
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *       &lt;sequence>
 *         &lt;element name="customAttributes" type="{http://attribute.ellipse.mincom.com}ArrayOfAttribute" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "AbstractDTO", propOrder = {
    "customAttributes"
})
@XmlSeeAlso({
    SecurityClassServiceModifyRequestDTO.class,
    SecurityClassServiceModifyMethodsRequestDTO.class,
    SecurityClassServiceModifyAttributesRequestDTO.class,
    SecurityClassServiceRetrieveRequestDTO.class,
    SecurityClassServiceRetrieveClassesRequestDTO.class,
    SecurityClassServiceReadRequestDTO.class,
    AbstractReplyDTO.class
})
public class AbstractDTO {

    protected ArrayOfAttribute customAttributes;

    /**
     * Gets the value of the customAttributes property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfAttribute }
     *     
     */
    public ArrayOfAttribute getCustomAttributes() {
        return customAttributes;
    }

    /**
     * Sets the value of the customAttributes property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfAttribute }
     *     
     */
    public void setCustomAttributes(ArrayOfAttribute value) {
        this.customAttributes = value;
    }

}
