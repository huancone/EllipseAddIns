
package com.mincom.enterpriseservice.ellipse.securityclass;

import java.math.BigDecimal;
import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.AbstractReplyDTO;


/**
 * <p>Java class for SecurityClassServiceRetrieveClassesReplyDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="SecurityClassServiceRetrieveClassesReplyDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractReplyDTO">
 *       &lt;sequence>
 *         &lt;element name="accessLevel" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}accessLevel" minOccurs="0"/>
 *         &lt;element name="appDesc" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}appDesc" minOccurs="0"/>
 *         &lt;element name="appName" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}appName" minOccurs="0"/>
 *         &lt;element name="appType" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}appType" minOccurs="0"/>
 *         &lt;element name="className" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}className" minOccurs="0"/>
 *         &lt;element name="districtCode" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}districtCode" minOccurs="0"/>
 *         &lt;element name="maxAccessLevel" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}maxAccessLevel" minOccurs="0"/>
 *         &lt;element name="primaryFlag" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}primaryFlag" minOccurs="0"/>
 *         &lt;element name="profileName" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}profileName" minOccurs="0"/>
 *         &lt;element name="profileType" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}profileType" minOccurs="0"/>
 *         &lt;element name="refcodeEntity" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}refcodeEntity" minOccurs="0"/>
 *         &lt;element name="reviewFlag" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}reviewFlag" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/extension>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "SecurityClassServiceRetrieveClassesReplyDTO", propOrder = {
    "accessLevel",
    "appDesc",
    "appName",
    "appType",
    "className",
    "districtCode",
    "maxAccessLevel",
    "primaryFlag",
    "profileName",
    "profileType",
    "refcodeEntity",
    "reviewFlag"
})
public class SecurityClassServiceRetrieveClassesReplyDTO
    extends AbstractReplyDTO
{

    protected BigDecimal accessLevel;
    protected String appDesc;
    protected String appName;
    protected String appType;
    protected String className;
    protected String districtCode;
    protected BigDecimal maxAccessLevel;
    protected Boolean primaryFlag;
    protected String profileName;
    protected String profileType;
    protected String refcodeEntity;
    protected Boolean reviewFlag;

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
     * Gets the value of the appDesc property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getAppDesc() {
        return appDesc;
    }

    /**
     * Sets the value of the appDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setAppDesc(String value) {
        this.appDesc = value;
    }

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
     * Gets the value of the maxAccessLevel property.
     * 
     * @return
     *     possible object is
     *     {@link BigDecimal }
     *     
     */
    public BigDecimal getMaxAccessLevel() {
        return maxAccessLevel;
    }

    /**
     * Sets the value of the maxAccessLevel property.
     * 
     * @param value
     *     allowed object is
     *     {@link BigDecimal }
     *     
     */
    public void setMaxAccessLevel(BigDecimal value) {
        this.maxAccessLevel = value;
    }

    /**
     * Gets the value of the primaryFlag property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isPrimaryFlag() {
        return primaryFlag;
    }

    /**
     * Sets the value of the primaryFlag property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setPrimaryFlag(Boolean value) {
        this.primaryFlag = value;
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
     * Gets the value of the refcodeEntity property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getRefcodeEntity() {
        return refcodeEntity;
    }

    /**
     * Sets the value of the refcodeEntity property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setRefcodeEntity(String value) {
        this.refcodeEntity = value;
    }

    /**
     * Gets the value of the reviewFlag property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReviewFlag() {
        return reviewFlag;
    }

    /**
     * Sets the value of the reviewFlag property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReviewFlag(Boolean value) {
        this.reviewFlag = value;
    }

}
