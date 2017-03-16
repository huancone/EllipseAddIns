
package com.mincom.enterpriseservice.ellipse.securityclass;

import java.util.ArrayList;
import java.util.List;
import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlType;


/**
 * <p>Java class for ArrayOfSecurityClassServiceModifyReplyDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="ArrayOfSecurityClassServiceModifyReplyDTO">
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *       &lt;sequence>
 *         &lt;element name="SecurityClassServiceModifyReplyDTO" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}SecurityClassServiceModifyReplyDTO" maxOccurs="unbounded" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "ArrayOfSecurityClassServiceModifyReplyDTO", propOrder = {
    "securityClassServiceModifyReplyDTO"
})
public class ArrayOfSecurityClassServiceModifyReplyDTO {

    @XmlElement(name = "SecurityClassServiceModifyReplyDTO", nillable = true)
    protected List<SecurityClassServiceModifyReplyDTO> securityClassServiceModifyReplyDTO;

    /**
     * Gets the value of the securityClassServiceModifyReplyDTO property.
     * 
     * <p>
     * This accessor method returns a reference to the live list,
     * not a snapshot. Therefore any modification you make to the
     * returned list will be present inside the JAXB object.
     * This is why there is not a <CODE>set</CODE> method for the securityClassServiceModifyReplyDTO property.
     * 
     * <p>
     * For example, to add a new item, do as follows:
     * <pre>
     *    getSecurityClassServiceModifyReplyDTO().add(newItem);
     * </pre>
     * 
     * 
     * <p>
     * Objects of the following type(s) are allowed in the list
     * {@link SecurityClassServiceModifyReplyDTO }
     * 
     * 
     */
    public List<SecurityClassServiceModifyReplyDTO> getSecurityClassServiceModifyReplyDTO() {
        if (securityClassServiceModifyReplyDTO == null) {
            securityClassServiceModifyReplyDTO = new ArrayList<SecurityClassServiceModifyReplyDTO>();
        }
        return this.securityClassServiceModifyReplyDTO;
    }

}
