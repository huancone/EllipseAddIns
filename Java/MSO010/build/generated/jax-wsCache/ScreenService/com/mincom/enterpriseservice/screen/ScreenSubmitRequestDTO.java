
package com.mincom.enterpriseservice.screen;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;


/**
 * <p>Java class for ScreenSubmitRequestDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="ScreenSubmitRequestDTO">
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *       &lt;sequence>
 *         &lt;element name="screenFields" type="{http://screen.enterpriseservice.mincom.com}ArrayOfScreenNameValueDTO" minOccurs="0"/>
 *         &lt;element name="screenKey" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "ScreenSubmitRequestDTO", propOrder = {
    "screenFields",
    "screenKey"
})
public class ScreenSubmitRequestDTO {

    protected ArrayOfScreenNameValueDTO screenFields;
    protected String screenKey;

    /**
     * Gets the value of the screenFields property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfScreenNameValueDTO }
     *     
     */
    public ArrayOfScreenNameValueDTO getScreenFields() {
        return screenFields;
    }

    /**
     * Sets the value of the screenFields property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfScreenNameValueDTO }
     *     
     */
    public void setScreenFields(ArrayOfScreenNameValueDTO value) {
        this.screenFields = value;
    }

    /**
     * Gets the value of the screenKey property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getScreenKey() {
        return screenKey;
    }

    /**
     * Sets the value of the screenKey property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setScreenKey(String value) {
        this.screenKey = value;
    }

}
