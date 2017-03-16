
package com.mincom.enterpriseservice.ellipse.securityclass;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.AbstractRequiredAttributesDTO;


/**
 * <p>Java class for SecurityClassServiceReadRequiredAttributesDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="SecurityClassServiceReadRequiredAttributesDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractRequiredAttributesDTO">
 *       &lt;sequence>
 *         &lt;element name="returnAccessLevel" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnClassName" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnDistrictCode" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnNoSecurityFlag" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnProfileName" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnProfileType" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/extension>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "SecurityClassServiceReadRequiredAttributesDTO", propOrder = {
    "returnAccessLevel",
    "returnClassName",
    "returnDistrictCode",
    "returnNoSecurityFlag",
    "returnProfileName",
    "returnProfileType"
})
public class SecurityClassServiceReadRequiredAttributesDTO
    extends AbstractRequiredAttributesDTO
{

    protected Boolean returnAccessLevel;
    protected Boolean returnClassName;
    protected Boolean returnDistrictCode;
    protected Boolean returnNoSecurityFlag;
    protected Boolean returnProfileName;
    protected Boolean returnProfileType;

    /**
     * Gets the value of the returnAccessLevel property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnAccessLevel() {
        return returnAccessLevel;
    }

    /**
     * Sets the value of the returnAccessLevel property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnAccessLevel(Boolean value) {
        this.returnAccessLevel = value;
    }

    /**
     * Gets the value of the returnClassName property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnClassName() {
        return returnClassName;
    }

    /**
     * Sets the value of the returnClassName property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnClassName(Boolean value) {
        this.returnClassName = value;
    }

    /**
     * Gets the value of the returnDistrictCode property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnDistrictCode() {
        return returnDistrictCode;
    }

    /**
     * Sets the value of the returnDistrictCode property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnDistrictCode(Boolean value) {
        this.returnDistrictCode = value;
    }

    /**
     * Gets the value of the returnNoSecurityFlag property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnNoSecurityFlag() {
        return returnNoSecurityFlag;
    }

    /**
     * Sets the value of the returnNoSecurityFlag property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnNoSecurityFlag(Boolean value) {
        this.returnNoSecurityFlag = value;
    }

    /**
     * Gets the value of the returnProfileName property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnProfileName() {
        return returnProfileName;
    }

    /**
     * Sets the value of the returnProfileName property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnProfileName(Boolean value) {
        this.returnProfileName = value;
    }

    /**
     * Gets the value of the returnProfileType property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnProfileType() {
        return returnProfileType;
    }

    /**
     * Sets the value of the returnProfileType property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnProfileType(Boolean value) {
        this.returnProfileType = value;
    }

}
