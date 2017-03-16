
package com.mincom.enterpriseservice.screen;

import java.util.ArrayList;
import java.util.List;
import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlType;


/**
 * <p>Java class for ArrayOfScreenFieldDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="ArrayOfScreenFieldDTO">
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *       &lt;sequence>
 *         &lt;element name="ScreenFieldDTO" type="{http://screen.enterpriseservice.mincom.com}ScreenFieldDTO" maxOccurs="unbounded" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "ArrayOfScreenFieldDTO", propOrder = {
    "screenFieldDTO"
})
public class ArrayOfScreenFieldDTO {

    @XmlElement(name = "ScreenFieldDTO", nillable = true)
    protected List<ScreenFieldDTO> screenFieldDTO;

    /**
     * Gets the value of the screenFieldDTO property.
     * 
     * <p>
     * This accessor method returns a reference to the live list,
     * not a snapshot. Therefore any modification you make to the
     * returned list will be present inside the JAXB object.
     * This is why there is not a <CODE>set</CODE> method for the screenFieldDTO property.
     * 
     * <p>
     * For example, to add a new item, do as follows:
     * <pre>
     *    getScreenFieldDTO().add(newItem);
     * </pre>
     * 
     * 
     * <p>
     * Objects of the following type(s) are allowed in the list
     * {@link ScreenFieldDTO }
     * 
     * 
     */
    public List<ScreenFieldDTO> getScreenFieldDTO() {
        if (screenFieldDTO == null) {
            screenFieldDTO = new ArrayList<ScreenFieldDTO>();
        }
        return this.screenFieldDTO;
    }

}
