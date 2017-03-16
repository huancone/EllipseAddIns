
package com.mincom.enterpriseservice.ellipse;

import java.util.ArrayList;
import java.util.List;
import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlType;


/**
 * <p>Java class for ArrayOfWarningMessageDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="ArrayOfWarningMessageDTO">
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *       &lt;sequence>
 *         &lt;element name="WarningMessageDTO" type="{http://ellipse.enterpriseservice.mincom.com}WarningMessageDTO" maxOccurs="unbounded" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "ArrayOfWarningMessageDTO", propOrder = {
    "warningMessageDTO"
})
public class ArrayOfWarningMessageDTO {

    @XmlElement(name = "WarningMessageDTO", nillable = true)
    protected List<WarningMessageDTO> warningMessageDTO;

    /**
     * Gets the value of the warningMessageDTO property.
     * 
     * <p>
     * This accessor method returns a reference to the live list,
     * not a snapshot. Therefore any modification you make to the
     * returned list will be present inside the JAXB object.
     * This is why there is not a <CODE>set</CODE> method for the warningMessageDTO property.
     * 
     * <p>
     * For example, to add a new item, do as follows:
     * <pre>
     *    getWarningMessageDTO().add(newItem);
     * </pre>
     * 
     * 
     * <p>
     * Objects of the following type(s) are allowed in the list
     * {@link WarningMessageDTO }
     * 
     * 
     */
    public List<WarningMessageDTO> getWarningMessageDTO() {
        if (warningMessageDTO == null) {
            warningMessageDTO = new ArrayList<WarningMessageDTO>();
        }
        return this.warningMessageDTO;
    }

}
