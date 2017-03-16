
package com.mincom.enterpriseservice.ellipse.securityclass;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.AbstractDTO;


/**
 * <p>Java class for SecurityClassServiceRetrieveClassesRequestDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="SecurityClassServiceRetrieveClassesRequestDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractDTO">
 *       &lt;sequence>
 *         &lt;element name="appName" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}appName" minOccurs="0"/>
 *         &lt;element name="appType" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}appType" minOccurs="0"/>
 *         &lt;element name="districtCode" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}districtCode" minOccurs="0"/>
 *         &lt;element name="profileName" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}profileName" minOccurs="0"/>
 *         &lt;element name="profileType" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}profileType" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/extension>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "SecurityClassServiceRetrieveClassesRequestDTO", propOrder = {
    "appName",
    "appType",
    "districtCode",
    "profileName",
    "profileType"
})
public class SecurityClassServiceRetrieveClassesRequestDTO
    extends AbstractDTO
{

    protected String appName;
    protected String appType;
    protected String districtCode;
    protected String profileName;
    protected String profileType;

    /**
     * Gets the value of the appName property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getAppName() {
        return appName;
    }

    /**
     * Sets the value of the appName property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setAppName(String value) {
        this.appName = value;
    }

    /**
     * Gets the value of the appType property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getAppType() {
        return appType;
    }

    /**
     * Sets the value of the appType property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setAppType(String value) {
        this.appType = value;
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

}
