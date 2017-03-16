
package com.mincom.enterpriseservice.ellipse.securityclass;

import java.util.ArrayList;
import java.util.List;
import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlType;


/**
 * <p>Java class for ArrayOfSecurityClassServiceRetrieveClassesReplyDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="ArrayOfSecurityClassServiceRetrieveClassesReplyDTO">
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *       &lt;sequence>
 *         &lt;element name="SecurityClassServiceRetrieveClassesReplyDTO" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}SecurityClassServiceRetrieveClassesReplyDTO" maxOccurs="unbounded" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "ArrayOfSecurityClassServiceRetrieveClassesReplyDTO", propOrder = {
    "securityClassServiceRetrieveClassesReplyDTO"
})
public class ArrayOfSecurityClassServiceRetrieveClassesReplyDTO {

    @XmlElement(name = "SecurityClassServiceRetrieveClassesReplyDTO", nillable = true)
    protected List<SecurityClassServiceRetrieveClassesReplyDTO> securityClassServiceRetrieveClassesReplyDTO;

    /**
     * Gets the value of the securityClassServiceRetrieveClassesReplyDTO property.
     * 
     * <p>
     * This accessor method returns a reference to the live list,
     * not a snapshot. Therefore any modification you make to the
     * returned list will be present inside the JAXB object.
     * This is why there is not a <CODE>set</CODE> method for the securityClassServiceRetrieveClassesReplyDTO property.
     * 
     * <p>
     * For example, to add a new item, do as follows:
     * <pre>
     *    getSecurityClassServiceRetrieveClassesReplyDTO().add(newItem);
     * </pre>
     * 
     * 
     * <p>
     * Objects of the following type(s) are allowed in the list
     * {@link SecurityClassServiceRetrieveClassesReplyDTO }
     * 
     * 
     */
    public List<SecurityClassServiceRetrieveClassesReplyDTO> getSecurityClassServiceRetrieveClassesReplyDTO() {
        if (securityClassServiceRetrieveClassesReplyDTO == null) {
            securityClassServiceRetrieveClassesReplyDTO = new ArrayList<SecurityClassServiceRetrieveClassesReplyDTO>();
        }
        return this.securityClassServiceRetrieveClassesReplyDTO;
    }

}
