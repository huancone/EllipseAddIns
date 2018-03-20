
package com.mincom.enterpriseservice.ellipse.securityclass;

import java.math.BigDecimal;
import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.AbstractDTO;


/**
 * <p>Java class for SecurityClassServiceModifyRequestDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="SecurityClassServiceModifyRequestDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractDTO">
 *       &lt;sequence>
 *         &lt;element name="accessLevel" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}accessLevel" minOccurs="0"/>
 *         &lt;element name="className" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}className" minOccurs="0"/>
 *         &lt;element name="districtCode" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}districtCode" minOccurs="0"/>
 *         &lt;element name="profileName" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}profileName" minOccurs="0"/>
 *         &lt;element name="profileType" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}profileType" minOccurs="0"/>
 *         &lt;element name="requiredAttributes" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}SecurityClassServiceModifyRequiredAttributesDTO" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/extension>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "SecurityClassServiceModifyRequestDTO", propOrder = {
    "accessLevel",
    "className",
    "districtCode",
    "profileName",
    "profileType",
    "requiredAttributes"
})
public class SecurityClassServiceModifyRequestDTO
    extends AbstractDTO
{

    protected BigDecimal accessLevel;
    protected String className;
    protected String districtCode;
    protected String profileName;
    protected String profileType;
    protected SecurityClassServiceModifyRequiredAttributesDTO requiredAttributes;

    /**
     * Gets the value of the accessLevel property.
     * 
     * @return
     *     possible object is
     *     {@link BigDecimal }
     *     
     */
    public BigDecimal getAccessLevel() {
        return accessLevel;
    }

    /**
     * Sets the value of the accessLevel property.
     * 
     * @param value
     *     allowed object is
     *     {@link BigDecimal }
     *     
     */
    public void setAccessLevel(BigDecimal value) {
        this.accessLevel = value;
    }

    /**
     * Gets the value of the className property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getClassName() {
        return className;
    }

    /**
     * Sets the value of the className property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setClassName(String value) {
        this.className = value;
    }

    /**
     * Gets the value of the districtCode property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getDistrictCode() {
        return districtCode;
    }

    /**
     * Sets the value of the districtCode property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setDistrictCode(String value) {
        this.districtCode = value;
    }

    /**
     * Gets the value of the profileName property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getProfileName() {
        return profileName;
    }

    /**
     * Sets the value of the profileName property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setProfileName(String value) {
        this.profileName = value;
    }

    /**
     * Gets the value of the profileType property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getProfileType() {
        return profileType;
    }

    /**
     * Sets the value of the profileType property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setProfileType(String value) {
        this.profileType = value;
    }

    /**
     * Gets the value of the requiredAttributes property.
     * 
     * @return
     *     possible object is
     *     {@link SecurityClassServiceModifyRequiredAttributesDTO }
     *     
     */
    public SecurityClassServiceModifyRequiredAttributesDTO getRequiredAttributes() {
        return requiredAttributes;
    }

    /**
     * Sets the value of the requiredAttributes property.
     * 
     * @param value
     *     allowed object is
     *     {@link SecurityClassServiceModifyRequiredAttributesDTO }
     *     
     */
    public void setRequiredAttributes(SecurityClassServiceModifyRequiredAttributesDTO value) {
        this.requiredAttributes = value;
    }

}
