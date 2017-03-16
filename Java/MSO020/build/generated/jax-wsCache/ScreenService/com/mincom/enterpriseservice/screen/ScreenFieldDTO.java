
package com.mincom.enterpriseservice.screen;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;


/**
 * <p>Java class for ScreenFieldDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="ScreenFieldDTO">
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *       &lt;sequence>
 *         &lt;element name="fieldName" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *         &lt;element name="fieldProtected" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="maximumLength" type="{http://www.w3.org/2001/XMLSchema}int" minOccurs="0"/>
 *         &lt;element name="value" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "ScreenFieldDTO", propOrder = {
    "fieldName",
    "fieldProtected",
    "maximumLength",
    "value"
})
public class ScreenFieldDTO {

    protected String fieldName;
    protected Boolean fieldProtected;
    protected Integer maximumLength;
    protected String value;

    /**
     * Gets the value of the fieldName property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getFieldName() {
        return fieldName;
    }

    /**
     * Sets the value of the fieldName property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setFieldName(String value) {
        this.fieldName = value;
    }

    /**
     * Gets the value of the fieldProtected property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isFieldProtected() {
        return fieldProtected;
    }

    /**
     * Sets the value of the fieldProtected property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setFieldProtected(Boolean value) {
        this.fieldProtected = value;
    }

    /**
     * Gets the value of the maximumLength property.
     * 
     * @return
     *     possible object is
     *     {@link Integer }
     *     
     */
    public Integer getMaximumLength() {
        return maximumLength;
    }

    /**
     * Sets the value of the maximumLength property.
     * 
     * @param value
     *     allowed object is
     *     {@link Integer }
     *     
     */
    public void setMaximumLength(Integer value) {
        this.maximumLength = value;
    }

    /**
     * Gets the value of the value property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getValue() {
        return value;
    }

    /**
     * Sets the value of the value property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setValue(String value) {
        this.value = value;
    }

}
