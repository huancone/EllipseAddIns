
package com.mincom.enterpriseservice.ellipse.securityclass;

import java.util.ArrayList;
import java.util.List;
import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlType;


/**
 * <p>Java class for ArrayOfSecurityClassServiceModifyMethodsRequestDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="ArrayOfSecurityClassServiceModifyMethodsRequestDTO">
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *       &lt;sequence>
 *         &lt;element name="SecurityClassServiceModifyMethodsRequestDTO" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}SecurityClassServiceModifyMethodsRequestDTO" maxOccurs="unbounded" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "ArrayOfSecurityClassServiceModifyMethodsRequestDTO", propOrder = {
    "securityClassServiceModifyMethodsRequestDTO"
})
public class ArrayOfSecurityClassServiceModifyMethodsRequestDTO {

    @XmlElement(name = "SecurityClassServiceModifyMethodsRequestDTO", nillable = true)
    protected List<SecurityClassServiceModifyMethodsRequestDTO> securityClassServiceModifyMethodsRequestDTO;

    /**
     * Gets the value of the securityClassServiceModifyMethodsRequestDTO property.
     * 
     * <p>
     * This accessor method returns a reference to the live list,
     * not a snapshot. Therefore any modification you make to the
     * returned list will be present inside the JAXB object.
     * This is why there is not a <CODE>set</CODE> method for the securityClassServiceModifyMethodsRequestDTO property.
     * 
     * <p>
     * For example, to add a new item, do as follows:
     * <pre>
     *    getSecurityClassServiceModifyMethodsRequestDTO().add(newItem);
     * </pre>
     * 
     * 
     * <p>
     * Objects of the following type(s) are allowed in the list
     * {@link SecurityClassServiceModifyMethodsRequestDTO }
     * 
     * 
     */
    public List<SecurityClassServiceModifyMethodsRequestDTO> getSecurityClassServiceModifyMethodsRequestDTO() {
        if (securityClassServiceModifyMethodsRequestDTO == null) {
            securityClassServiceModifyMethodsRequestDTO = new ArrayList<SecurityClassServiceModifyMethodsRequestDTO>();
        }
        return this.securityClassServiceModifyMethodsRequestDTO;
    }

}
