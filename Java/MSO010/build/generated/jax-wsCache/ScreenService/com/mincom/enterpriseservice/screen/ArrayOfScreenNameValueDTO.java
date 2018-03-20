
package com.mincom.enterpriseservice.screen;

import java.util.ArrayList;
import java.util.List;
import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlType;


/**
 * <p>Java class for ArrayOfScreenNameValueDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="ArrayOfScreenNameValueDTO">
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *       &lt;sequence>
 *         &lt;element name="ScreenNameValueDTO" type="{http://screen.enterpriseservice.mincom.com}ScreenNameValueDTO" maxOccurs="unbounded" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "ArrayOfScreenNameValueDTO", propOrder = {
    "screenNameValueDTO"
})
public class ArrayOfScreenNameValueDTO {

    @XmlElement(name = "ScreenNameValueDTO", nillable = true)
    protected List<ScreenNameValueDTO> screenNameValueDTO;

    /**
     * Gets the value of the screenNameValueDTO property.
     * 
     * <p>
     * This accessor method returns a reference to the live list,
     * not a snapshot. Therefore any modification you make to the
     * returned list will be present inside the JAXB object.
     * This is why there is not a <CODE>set</CODE> method for the screenNameValueDTO property.
     * 
     * <p>
     * For example, to add a new item, do as follows:
     * <pre>
     *    getScreenNameValueDTO().add(newItem);
     * </pre>
     * 
     * 
     * <p>
     * Objects of the following type(s) are allowed in the list
     * {@link ScreenNameValueDTO }
     * 
     * 
     */
    public List<ScreenNameValueDTO> getScreenNameValueDTO() {
        if (screenNameValueDTO == null) {
            screenNameValueDTO = new ArrayList<ScreenNameValueDTO>();
        }
        return this.screenNameValueDTO;
    }

}
