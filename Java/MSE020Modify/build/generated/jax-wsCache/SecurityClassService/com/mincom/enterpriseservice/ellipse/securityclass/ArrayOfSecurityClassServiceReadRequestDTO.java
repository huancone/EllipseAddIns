
package com.mincom.enterpriseservice.ellipse.securityclass;

import java.util.ArrayList;
import java.util.List;
import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlType;


/**
 * <p>Java class for ArrayOfSecurityClassServiceReadRequestDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="ArrayOfSecurityClassServiceReadRequestDTO">
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *       &lt;sequence>
 *         &lt;element name="SecurityClassServiceReadRequestDTO" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}SecurityClassServiceReadRequestDTO" maxOccurs="unbounded" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "ArrayOfSecurityClassServiceReadRequestDTO", propOrder = {
    "securityClassServiceReadRequestDTO"
})
public class ArrayOfSecurityClassServiceReadRequestDTO {

    @XmlElement(name = "SecurityClassServiceReadRequestDTO", nillable = true)
    protected List<SecurityClassServiceReadRequestDTO> securityClassServiceReadRequestDTO;

    /**
     * Gets the value of the securityClassServiceReadRequestDTO property.
     * 
     * <p>
     * This accessor method returns a reference to the live list,
     * not a snapshot. Therefore any modification you make to the
     * returned list will be present inside the JAXB object.
     * This is why there is not a <CODE>set</CODE> method for the securityClassServiceReadRequestDTO property.
     * 
     * <p>
     * For example, to add a new item, do as follows:
     * <pre>
     *    getSecurityClassServiceReadRequestDTO().add(newItem);
     * </pre>
     * 
     * 
     * <p>
     * Objects of the following type(s) are allowed in the list
     * {@link SecurityClassServiceReadRequestDTO }
     * 
     * 
     */
    public List<SecurityClassServiceReadRequestDTO> getSecurityClassServiceReadRequestDTO() {
        if (securityClassServiceReadRequestDTO == null) {
            securityClassServiceReadRequestDTO = new ArrayList<SecurityClassServiceReadRequestDTO>();
        }
        return this.securityClassServiceReadRequestDTO;
    }

}
