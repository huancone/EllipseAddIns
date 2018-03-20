
package com.mincom.enterpriseservice.ellipse.securityclass;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.AbstractRequiredAttributesDTO;


/**
 * <p>Java class for SecurityClassServiceRetrieveClassesRequiredAttributesDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="SecurityClassServiceRetrieveClassesRequiredAttributesDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractRequiredAttributesDTO">
 *       &lt;sequence>
 *         &lt;element name="returnAccessLevel" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnAppDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnAppName" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnAppType" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnClassName" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnDistrictCode" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnMaxAccessLevel" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPrimaryFlag" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnProfileName" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnProfileType" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnRefcodeEntity" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnReviewFlag" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/extension>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "SecurityClassServiceRetrieveClassesRequiredAttributesDTO", propOrder = {
    "returnAccessLevel",
    "returnAppDesc",
    "returnAppName",
    "returnAppType",
    "returnClassName",
    "returnDistrictCode",
    "returnMaxAccessLevel",
    "returnPrimaryFlag",
    "returnProfileName",
    "returnProfileType",
    "returnRefcodeEntity",
    "returnReviewFlag"
})
public class SecurityClassServiceRetrieveClassesRequiredAttributesDTO
    extends AbstractRequiredAttributesDTO
{

    protected Boolean returnAccessLevel;
    protected Boolean returnAppDesc;
    protected Boolean returnAppName;
    protected Boolean returnAppType;
    protected Boolean returnClassName;
    protected Boolean returnDistrictCode;
    protected Boolean returnMaxAccessLevel;
    protected Boolean returnPrimaryFlag;
    protected Boolean returnProfileName;
    protected Boolean returnProfileType;
    protected Boolean returnRefcodeEntity;
    protected Boolean returnReviewFlag;

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
     * Gets the value of the returnAppDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnAppDesc() {
        return returnAppDesc;
    }

    /**
     * Sets the value of the returnAppDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnAppDesc(Boolean value) {
        this.returnAppDesc = value;
    }

    /**
     * Gets the value of the returnAppName property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnAppName() {
        return returnAppName;
    }

    /**
     * Sets the value of the returnAppName property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnAppName(Boolean value) {
        this.returnAppName = value;
    }

    /**
     * Gets the value of the returnAppType property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnAppType() {
        return returnAppType;
    }

    /**
     * Sets the value of the returnAppType property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnAppType(Boolean value) {
        this.returnAppType = value;
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
     * Gets the value of the returnMaxAccessLevel property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnMaxAccessLevel() {
        return returnMaxAccessLevel;
    }

    /**
     * Sets the value of the returnMaxAccessLevel property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnMaxAccessLevel(Boolean value) {
        this.returnMaxAccessLevel = value;
    }

    /**
     * Gets the value of the returnPrimaryFlag property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPrimaryFlag() {
        return returnPrimaryFlag;
    }

    /**
     * Sets the value of the returnPrimaryFlag property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPrimaryFlag(Boolean value) {
        this.returnPrimaryFlag = value;
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

    /**
     * Gets the value of the returnRefcodeEntity property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnRefcodeEntity() {
        return returnRefcodeEntity;
    }

    /**
     * Sets the value of the returnRefcodeEntity property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnRefcodeEntity(Boolean value) {
        this.returnRefcodeEntity = value;
    }

    /**
     * Gets the value of the returnReviewFlag property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnReviewFlag() {
        return returnReviewFlag;
    }

    /**
     * Sets the value of the returnReviewFlag property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnReviewFlag(Boolean value) {
        this.returnReviewFlag = value;
    }

}
