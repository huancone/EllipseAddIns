
package com.mincom.enterpriseservice.ellipse.securityclass;

import java.util.ArrayList;
import java.util.List;
import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlType;


/**
 * <p>Java class for ArrayOfSecurityClassServiceModifyMethodsReplyDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="ArrayOfSecurityClassServiceModifyMethodsReplyDTO">
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *       &lt;sequence>
 *         &lt;element name="SecurityClassServiceModifyMethodsReplyDTO" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}SecurityClassServiceModifyMethodsReplyDTO" maxOccurs="unbounded" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "ArrayOfSecurityClassServiceModifyMethodsReplyDTO", propOrder = {
    "securityClassServiceModifyMethodsReplyDTO"
})
public class ArrayOfSecurityClassServiceModifyMethodsReplyDTO {

    @XmlElement(name = "SecurityClassServiceModifyMethodsReplyDTO", nillable = true)
    protected List<SecurityClassServiceModifyMethodsReplyDTO> securityClassServiceModifyMethodsReplyDTO;

    /**
     * Gets the value of the securityClassServiceModifyMethodsReplyDTO property.
     * 
     * <p>
     * This accessor method returns a reference to the live list,
     * not a snapshot. Therefore any modification you make to the
     * returned list will be present inside the JAXB object.
     * This is why there is not a <CODE>set</CODE> method for the securityClassServiceModifyMethodsReplyDTO property.
     * 
     * <p>
     * For example, to add a new item, do as follows:
     * <pre>
     *    getSecurityClassServiceModifyMethodsReplyDTO().add(newItem);
     * </pre>
     * 
     * 
     * <p>
     * Objects of the following type(s) are allowed in the list
     * {@link SecurityClassServiceModifyMethodsReplyDTO }
     * 
     * 
     */
    public List<SecurityClassServiceModifyMethodsReplyDTO> getSecurityClassServiceModifyMethodsReplyDTO() {
        if (securityClassServiceModifyMethodsReplyDTO == null) {
            securityClassServiceModifyMethodsReplyDTO = new ArrayList<SecurityClassServiceModifyMethodsReplyDTO>();
        }
        return this.securityClassServiceModifyMethodsReplyDTO;
    }

}
