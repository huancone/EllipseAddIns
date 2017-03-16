
package com.mincom.enterpriseservice.ellipse.employee;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.AbstractReplyDTO;


/**
 * <p>Java class for EmployeeServiceShowReplyDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="EmployeeServiceShowReplyDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractReplyDTO">
 *       &lt;sequence>
 *         &lt;element name="employee" type="{http://employee.ellipse.enterpriseservice.mincom.com}employee" minOccurs="0"/>
 *         &lt;element name="employeeFormattedName" type="{http://employee.ellipse.enterpriseservice.mincom.com}employeeFormattedName" minOccurs="0"/>
 *         &lt;element name="languageCodeDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}languageCodeDesc" minOccurs="0"/>
 *         &lt;element name="messagePreferenceDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}messagePreferenceDesc" minOccurs="0"/>
 *         &lt;element name="postalCountryDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}postalCountryDesc" minOccurs="0"/>
 *         &lt;element name="postalStateDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}postalStateDesc" minOccurs="0"/>
 *         &lt;element name="printerCode1" type="{http://employee.ellipse.enterpriseservice.mincom.com}printerCode1" minOccurs="0"/>
 *         &lt;element name="printerDesc1" type="{http://employee.ellipse.enterpriseservice.mincom.com}printerDesc1" minOccurs="0"/>
 *         &lt;element name="residentialCountryDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}residentialCountryDesc" minOccurs="0"/>
 *         &lt;element name="residentialStateDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}residentialStateDesc" minOccurs="0"/>
 *         &lt;element name="siteCodeDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}siteCodeDesc" minOccurs="0"/>
 *         &lt;element name="socialInsuranceNumber" type="{http://employee.ellipse.enterpriseservice.mincom.com}socialInsuranceNumber" minOccurs="0"/>
 *         &lt;element name="socialSecurityNoTypeDescription" type="{http://employee.ellipse.enterpriseservice.mincom.com}socialSecurityNoTypeDescription" minOccurs="0"/>
 *         &lt;element name="socialSecurityNumber" type="{http://employee.ellipse.enterpriseservice.mincom.com}socialSecurityNumber" minOccurs="0"/>
 *         &lt;element name="titleDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}titleDesc" minOccurs="0"/>
 *         &lt;element name="workOrderPrefixDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}workOrderPrefixDesc" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/extension>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "EmployeeServiceShowReplyDTO", propOrder = {
    "employee",
    "employeeFormattedName",
    "languageCodeDesc",
    "messagePreferenceDesc",
    "postalCountryDesc",
    "postalStateDesc",
    "printerCode1",
    "printerDesc1",
    "residentialCountryDesc",
    "residentialStateDesc",
    "siteCodeDesc",
    "socialInsuranceNumber",
    "socialSecurityNoTypeDescription",
    "socialSecurityNumber",
    "titleDesc",
    "workOrderPrefixDesc"
})
public class EmployeeServiceShowReplyDTO
    extends AbstractReplyDTO
{

    protected String employee;
    protected String employeeFormattedName;
    protected String languageCodeDesc;
    protected String messagePreferenceDesc;
    protected String postalCountryDesc;
    protected String postalStateDesc;
    protected String printerCode1;
    protected String printerDesc1;
    protected String residentialCountryDesc;
    protected String residentialStateDesc;
    protected String siteCodeDesc;
    protected String socialInsuranceNumber;
    protected String socialSecurityNoTypeDescription;
    protected String socialSecurityNumber;
    protected String titleDesc;
    protected String workOrderPrefixDesc;

    /**
     * Gets the value of the employee property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getEmployee() {
        return employee;
    }

    /**
     * Sets the value of the employee property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setEmployee(String value) {
        this.employee = value;
    }

    /**
     * Gets the value of the employeeFormattedName property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getEmployeeFormattedName() {
        return employeeFormattedName;
    }

    /**
     * Sets the value of the employeeFormattedName property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setEmployeeFormattedName(String value) {
        this.employeeFormattedName = value;
    }

    /**
     * Gets the value of the languageCodeDesc property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getLanguageCodeDesc() {
        return languageCodeDesc;
    }

    /**
     * Sets the value of the languageCodeDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setLanguageCodeDesc(String value) {
        this.languageCodeDesc = value;
    }

    /**
     * Gets the value of the messagePreferenceDesc property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getMessagePreferenceDesc() {
        return messagePreferenceDesc;
    }

    /**
     * Sets the value of the messagePreferenceDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setMessagePreferenceDesc(String value) {
        this.messagePreferenceDesc = value;
    }

    /**
     * Gets the value of the postalCountryDesc property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPostalCountryDesc() {
        return postalCountryDesc;
    }

    /**
     * Sets the value of the postalCountryDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPostalCountryDesc(String value) {
        this.postalCountryDesc = value;
    }

    /**
     * Gets the value of the postalStateDesc property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPostalStateDesc() {
        return postalStateDesc;
    }

    /**
     * Sets the value of the postalStateDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPostalStateDesc(String value) {
        this.postalStateDesc = value;
    }

    /**
     * Gets the value of the printerCode1 property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPrinterCode1() {
        return printerCode1;
    }

    /**
     * Sets the value of the printerCode1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPrinterCode1(String value) {
        this.printerCode1 = value;
    }

    /**
     * Gets the value of the printerDesc1 property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPrinterDesc1() {
        return printerDesc1;
    }

    /**
     * Sets the value of the printerDesc1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPrinterDesc1(String value) {
        this.printerDesc1 = value;
    }

    /**
     * Gets the value of the residentialCountryDesc property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getResidentialCountryDesc() {
        return residentialCountryDesc;
    }

    /**
     * Sets the value of the residentialCountryDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setResidentialCountryDesc(String value) {
        this.residentialCountryDesc = value;
    }

    /**
     * Gets the value of the residentialStateDesc property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getResidentialStateDesc() {
        return residentialStateDesc;
    }

    /**
     * Sets the value of the residentialStateDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setResidentialStateDesc(String value) {
        this.residentialStateDesc = value;
    }

    /**
     * Gets the value of the siteCodeDesc property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getSiteCodeDesc() {
        return siteCodeDesc;
    }

    /**
     * Sets the value of the siteCodeDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setSiteCodeDesc(String value) {
        this.siteCodeDesc = value;
    }

    /**
     * Gets the value of the socialInsuranceNumber property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getSocialInsuranceNumber() {
        return socialInsuranceNumber;
    }

    /**
     * Sets the value of the socialInsuranceNumber property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setSocialInsuranceNumber(String value) {
        this.socialInsuranceNumber = value;
    }

    /**
     * Gets the value of the socialSecurityNoTypeDescription property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getSocialSecurityNoTypeDescription() {
        return socialSecurityNoTypeDescription;
    }

    /**
     * Sets the value of the socialSecurityNoTypeDescription property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setSocialSecurityNoTypeDescription(String value) {
        this.socialSecurityNoTypeDescription = value;
    }

    /**
     * Gets the value of the socialSecurityNumber property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getSocialSecurityNumber() {
        return socialSecurityNumber;
    }

    /**
     * Sets the value of the socialSecurityNumber property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setSocialSecurityNumber(String value) {
        this.socialSecurityNumber = value;
    }

    /**
     * Gets the value of the titleDesc property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getTitleDesc() {
        return titleDesc;
    }

    /**
     * Sets the value of the titleDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setTitleDesc(String value) {
        this.titleDesc = value;
    }

    /**
     * Gets the value of the workOrderPrefixDesc property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getWorkOrderPrefixDesc() {
        return workOrderPrefixDesc;
    }

    /**
     * Sets the value of the workOrderPrefixDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setWorkOrderPrefixDesc(String value) {
        this.workOrderPrefixDesc = value;
    }

}
