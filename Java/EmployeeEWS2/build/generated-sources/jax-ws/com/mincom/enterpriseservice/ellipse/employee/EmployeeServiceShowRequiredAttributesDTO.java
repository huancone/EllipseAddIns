
package com.mincom.enterpriseservice.ellipse.employee;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.AbstractRequiredAttributesDTO;


/**
 * <p>Java class for EmployeeServiceShowRequiredAttributesDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="EmployeeServiceShowRequiredAttributesDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractRequiredAttributesDTO">
 *       &lt;sequence>
 *         &lt;element name="returnEmployee" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnEmployeeFormattedName" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnLanguageCodeDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnMessagePreferenceDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPostalCountryDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPostalStateDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPrinterCode1" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPrinterDesc1" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnResidentialCountryDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnResidentialStateDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSiteCodeDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSocialInsuranceNumber" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSocialSecurityNoTypeDescription" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSocialSecurityNumber" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnTitleDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnWorkOrderPrefixDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/extension>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "EmployeeServiceShowRequiredAttributesDTO", propOrder = {
    "returnEmployee",
    "returnEmployeeFormattedName",
    "returnLanguageCodeDesc",
    "returnMessagePreferenceDesc",
    "returnPostalCountryDesc",
    "returnPostalStateDesc",
    "returnPrinterCode1",
    "returnPrinterDesc1",
    "returnResidentialCountryDesc",
    "returnResidentialStateDesc",
    "returnSiteCodeDesc",
    "returnSocialInsuranceNumber",
    "returnSocialSecurityNoTypeDescription",
    "returnSocialSecurityNumber",
    "returnTitleDesc",
    "returnWorkOrderPrefixDesc"
})
public class EmployeeServiceShowRequiredAttributesDTO
    extends AbstractRequiredAttributesDTO
{

    protected Boolean returnEmployee;
    protected Boolean returnEmployeeFormattedName;
    protected Boolean returnLanguageCodeDesc;
    protected Boolean returnMessagePreferenceDesc;
    protected Boolean returnPostalCountryDesc;
    protected Boolean returnPostalStateDesc;
    protected Boolean returnPrinterCode1;
    protected Boolean returnPrinterDesc1;
    protected Boolean returnResidentialCountryDesc;
    protected Boolean returnResidentialStateDesc;
    protected Boolean returnSiteCodeDesc;
    protected Boolean returnSocialInsuranceNumber;
    protected Boolean returnSocialSecurityNoTypeDescription;
    protected Boolean returnSocialSecurityNumber;
    protected Boolean returnTitleDesc;
    protected Boolean returnWorkOrderPrefixDesc;

    /**
     * Gets the value of the returnEmployee property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnEmployee() {
        return returnEmployee;
    }

    /**
     * Sets the value of the returnEmployee property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnEmployee(Boolean value) {
        this.returnEmployee = value;
    }

    /**
     * Gets the value of the returnEmployeeFormattedName property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnEmployeeFormattedName() {
        return returnEmployeeFormattedName;
    }

    /**
     * Sets the value of the returnEmployeeFormattedName property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnEmployeeFormattedName(Boolean value) {
        this.returnEmployeeFormattedName = value;
    }

    /**
     * Gets the value of the returnLanguageCodeDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnLanguageCodeDesc() {
        return returnLanguageCodeDesc;
    }

    /**
     * Sets the value of the returnLanguageCodeDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnLanguageCodeDesc(Boolean value) {
        this.returnLanguageCodeDesc = value;
    }

    /**
     * Gets the value of the returnMessagePreferenceDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnMessagePreferenceDesc() {
        return returnMessagePreferenceDesc;
    }

    /**
     * Sets the value of the returnMessagePreferenceDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnMessagePreferenceDesc(Boolean value) {
        this.returnMessagePreferenceDesc = value;
    }

    /**
     * Gets the value of the returnPostalCountryDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPostalCountryDesc() {
        return returnPostalCountryDesc;
    }

    /**
     * Sets the value of the returnPostalCountryDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPostalCountryDesc(Boolean value) {
        this.returnPostalCountryDesc = value;
    }

    /**
     * Gets the value of the returnPostalStateDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPostalStateDesc() {
        return returnPostalStateDesc;
    }

    /**
     * Sets the value of the returnPostalStateDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPostalStateDesc(Boolean value) {
        this.returnPostalStateDesc = value;
    }

    /**
     * Gets the value of the returnPrinterCode1 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPrinterCode1() {
        return returnPrinterCode1;
    }

    /**
     * Sets the value of the returnPrinterCode1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPrinterCode1(Boolean value) {
        this.returnPrinterCode1 = value;
    }

    /**
     * Gets the value of the returnPrinterDesc1 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPrinterDesc1() {
        return returnPrinterDesc1;
    }

    /**
     * Sets the value of the returnPrinterDesc1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPrinterDesc1(Boolean value) {
        this.returnPrinterDesc1 = value;
    }

    /**
     * Gets the value of the returnResidentialCountryDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnResidentialCountryDesc() {
        return returnResidentialCountryDesc;
    }

    /**
     * Sets the value of the returnResidentialCountryDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnResidentialCountryDesc(Boolean value) {
        this.returnResidentialCountryDesc = value;
    }

    /**
     * Gets the value of the returnResidentialStateDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnResidentialStateDesc() {
        return returnResidentialStateDesc;
    }

    /**
     * Sets the value of the returnResidentialStateDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnResidentialStateDesc(Boolean value) {
        this.returnResidentialStateDesc = value;
    }

    /**
     * Gets the value of the returnSiteCodeDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSiteCodeDesc() {
        return returnSiteCodeDesc;
    }

    /**
     * Sets the value of the returnSiteCodeDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSiteCodeDesc(Boolean value) {
        this.returnSiteCodeDesc = value;
    }

    /**
     * Gets the value of the returnSocialInsuranceNumber property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSocialInsuranceNumber() {
        return returnSocialInsuranceNumber;
    }

    /**
     * Sets the value of the returnSocialInsuranceNumber property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSocialInsuranceNumber(Boolean value) {
        this.returnSocialInsuranceNumber = value;
    }

    /**
     * Gets the value of the returnSocialSecurityNoTypeDescription property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSocialSecurityNoTypeDescription() {
        return returnSocialSecurityNoTypeDescription;
    }

    /**
     * Sets the value of the returnSocialSecurityNoTypeDescription property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSocialSecurityNoTypeDescription(Boolean value) {
        this.returnSocialSecurityNoTypeDescription = value;
    }

    /**
     * Gets the value of the returnSocialSecurityNumber property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSocialSecurityNumber() {
        return returnSocialSecurityNumber;
    }

    /**
     * Sets the value of the returnSocialSecurityNumber property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSocialSecurityNumber(Boolean value) {
        this.returnSocialSecurityNumber = value;
    }

    /**
     * Gets the value of the returnTitleDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnTitleDesc() {
        return returnTitleDesc;
    }

    /**
     * Sets the value of the returnTitleDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnTitleDesc(Boolean value) {
        this.returnTitleDesc = value;
    }

    /**
     * Gets the value of the returnWorkOrderPrefixDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnWorkOrderPrefixDesc() {
        return returnWorkOrderPrefixDesc;
    }

    /**
     * Sets the value of the returnWorkOrderPrefixDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnWorkOrderPrefixDesc(Boolean value) {
        this.returnWorkOrderPrefixDesc = value;
    }

}
