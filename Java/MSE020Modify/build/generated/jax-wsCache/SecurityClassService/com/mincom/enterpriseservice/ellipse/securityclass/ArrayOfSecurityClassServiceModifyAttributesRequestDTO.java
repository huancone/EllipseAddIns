
package com.mincom.enterpriseservice.ellipse.securityclass;

import java.util.ArrayList;
import java.util.List;
import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlType;


/**
 * <p>Java class for ArrayOfSecurityClassServiceModifyAttributesRequestDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="ArrayOfSecurityClassServiceModifyAttributesRequestDTO">
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *       &lt;sequence>
 *         &lt;element name="SecurityClassServiceModifyAttributesRequestDTO" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}SecurityClassServiceModifyAttributesRequestDTO" maxOccurs="unbounded" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "ArrayOfSecurityClassServiceModifyAttributesRequestDTO", propOrder = {
    "securityClassServiceModifyAttributesRequestDTO"
})
public class ArrayOfSecurityClassServiceModifyAttributesRequestDTO {

    @XmlElement(name = "SecurityClassServiceModifyAttributesRequestDTO", nillable = true)
    protected List<SecurityClassServiceModifyAttributesRequestDTO> securityClassServiceModifyAttributesRequestDTO;

    /**
     * Gets the value of the securityClassServiceModifyAttributesRequestDTO property.
     * 
     * <p>
     * This accessor method returns a reference to the live list,
     * not a snapshot. Therefore any modification you make to the
     * returned list will be present inside the JAXB object.
     * This is why there is not a <CODE>set</CODE> method for the securityClassServiceModifyAttributesRequestDTO property.
     * 
     * <p>
     * For example, to add a new item, do as follows:
     * <pre>
     *    getSecurityClassServiceModifyAttributesRequestDTO().add(newItem);
     * </pre>
     * 
     * 
     * <p>
     * Objects of the following type(s) are allowed in the list
     * {@link SecurityClassServiceModifyAttributesRequestDTO }
     * 
     * 
     */
    public List<SecurityClassServiceModifyAttributesRequestDTO> getSecurityClassServiceModifyAttributesRequestDTO() {
        if (securityClassServiceModifyAttributesRequestDTO == null) {
            securityClassServiceModifyAttributesRequestDTO = new ArrayList<SecurityClassServiceModifyAttributesRequestDTO>();
        }
        return this.securityClassServiceModifyAttributesRequestDTO;
    }

}
