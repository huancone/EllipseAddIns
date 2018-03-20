
package com.mincom.enterpriseservice.ellipse.employee;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.AbstractDTO;


/**
 * <p>Java class for EmployeeServiceShowRequestDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="EmployeeServiceShowRequestDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractDTO">
 *       &lt;sequence>
 *         &lt;element name="employee" type="{http://employee.ellipse.enterpriseservice.mincom.com}employee" minOccurs="0"/>
 *         &lt;element name="firstName" type="{http://employee.ellipse.enterpriseservice.mincom.com}firstName" minOccurs="0"/>
 *         &lt;element name="languageCode" type="{http://employee.ellipse.enterpriseservice.mincom.com}languageCode" minOccurs="0"/>
 *         &lt;element name="lastName" type="{http://employee.ellipse.enterpriseservice.mincom.com}lastName" minOccurs="0"/>
 *         &lt;element name="messagePreference" type="{http://employee.ellipse.enterpriseservice.mincom.com}messagePreference" minOccurs="0"/>
 *         &lt;element name="postalCountry" type="{http://employee.ellipse.enterpriseservice.mincom.com}postalCountry" minOccurs="0"/>
 *         &lt;element name="postalState" type="{http://employee.ellipse.enterpriseservice.mincom.com}postalState" minOccurs="0"/>
 *         &lt;element name="preferredName" type="{http://employee.ellipse.enterpriseservice.mincom.com}preferredName" minOccurs="0"/>
 *         &lt;element name="printerName1" type="{http://employee.ellipse.enterpriseservice.mincom.com}printerName1" minOccurs="0"/>
 *         &lt;element name="requiredAttributes" type="{http://employee.ellipse.enterpriseservice.mincom.com}EmployeeServiceShowRequiredAttributesDTO" minOccurs="0"/>
 *         &lt;element name="residentialCountry" type="{http://employee.ellipse.enterpriseservice.mincom.com}residentialCountry" minOccurs="0"/>
 *         &lt;element name="residentialState" type="{http://employee.ellipse.enterpriseservice.mincom.com}residentialState" minOccurs="0"/>
 *         &lt;element name="secondName" type="{http://employee.ellipse.enterpriseservice.mincom.com}secondName" minOccurs="0"/>
 *         &lt;element name="siteCode" type="{http://employee.ellipse.enterpriseservice.mincom.com}siteCode" minOccurs="0"/>
 *         &lt;element name="socialInsuranceNumber" type="{http://employee.ellipse.enterpriseservice.mincom.com}socialInsuranceNumber" minOccurs="0"/>
 *         &lt;element name="socialSecurityNoType" type="{http://employee.ellipse.enterpriseservice.mincom.com}socialSecurityNoType" minOccurs="0"/>
 *         &lt;element name="socialSecurityNumber" type="{http://employee.ellipse.enterpriseservice.mincom.com}socialSecurityNumber" minOccurs="0"/>
 *         &lt;element name="thirdName" type="{http://employee.ellipse.enterpriseservice.mincom.com}thirdName" minOccurs="0"/>
 *         &lt;element name="title" type="{http://employee.ellipse.enterpriseservice.mincom.com}title" minOccurs="0"/>
 *         &lt;element name="workOrderPrefix" type="{http://employee.ellipse.enterpriseservice.mincom.com}workOrderPrefix" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/extension>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "EmployeeServiceShowRequestDTO", propOrder = {
    "employee",
    "firstName",
    "languageCode",
    "lastName",
    "messagePreference",
    "postalCountry",
    "postalState",
    "preferredName",
    "printerName1",
    "requiredAttributes",
    "residentialCountry",
    "residentialState",
    "secondName",
    "siteCode",
    "socialInsuranceNumber",
    "socialSecurityNoType",
    "socialSecurityNumber",
    "thirdName",
    "title",
    "workOrderPrefix"
})
public class EmployeeServiceShowRequestDTO
    extends AbstractDTO
{

    protected String employee;
    protected String firstName;
    protected String languageCode;
    protected String lastName;
    protected String messagePreference;
    protected String postalCountry;
    protected String postalState;
    protected String preferredName;
    protected String printerName1;
    protected EmployeeServiceShowRequiredAttributesDTO requiredAttributes;
    protected String residentialCountry;
    protected String residentialState;
    protected String secondName;
    protected String siteCode;
    protected String socialInsuranceNumber;
    protected String socialSecurityNoType;
    protected String socialSecurityNumber;
    protected String thirdName;
    protected String title;
    protected String workOrderPrefix;

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
     * Gets the value of the firstName property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getFirstName() {
        return firstName;
    }

    /**
     * Sets the value of the firstName property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setFirstName(String value) {
        this.firstName = value;
    }

    /**
     * Gets the value of the languageCode property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getLanguageCode() {
        return languageCode;
    }

    /**
     * Sets the value of the languageCode property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setLanguageCode(String value) {
        this.languageCode = value;
    }

    /**
     * Gets the value of the lastName property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getLastName() {
        return lastName;
    }

    /**
     * Sets the value of the lastName property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setLastName(String value) {
        this.lastName = value;
    }

    /**
     * Gets the value of the messagePreference property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getMessagePreference() {
        return messagePreference;
    }

    /**
     * Sets the value of the messagePreference property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setMessagePreference(String value) {
        this.messagePreference = value;
    }

    /**
     * Gets the value of the postalCountry property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPostalCountry() {
        return postalCountry;
    }

    /**
     * Sets the value of the postalCountry property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPostalCountry(String value) {
        this.postalCountry = value;
    }

    /**
     * Gets the value of the postalState property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPostalState() {
        return postalState;
    }

    /**
     * Sets the value of the postalState property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPostalState(String value) {
        this.postalState = value;
    }

    /**
     * Gets the value of the preferredName property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPreferredName() {
        return preferredName;
    }

    /**
     * Sets the value of the preferredName property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPreferredName(String value) {
        this.preferredName = value;
    }

    /**
     * Gets the value of the printerName1 property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPrinterName1() {
        return printerName1;
    }

    /**
     * Sets the value of the printerName1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPrinterName1(String value) {
        this.printerName1 = value;
    }

    /**
     * Gets the value of the requiredAttributes property.
     * 
     * @return
     *     possible object is
     *     {@link EmployeeServiceShowRequiredAttributesDTO }
     *     
     */
    public EmployeeServiceShowRequiredAttributesDTO getRequiredAttributes() {
        return requiredAttributes;
    }

    /**
     * Sets the value of the requiredAttributes property.
     * 
     * @param value
     *     allowed object is
     *     {@link EmployeeServiceShowRequiredAttributesDTO }
     *     
     */
    public void setRequiredAttributes(EmployeeServiceShowRequiredAttributesDTO value) {
        this.requiredAttributes = value;
    }

    /**
     * Gets the value of the residentialCountry property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getResidentialCountry() {
        return residentialCountry;
    }

    /**
     * Sets the value of the residentialCountry property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setResidentialCountry(String value) {
        this.residentialCountry = value;
    }

    /**
     * Gets the value of the residentialState property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getResidentialState() {
        return residentialState;
    }

    /**
     * Sets the value of the residentialState property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setResidentialState(String value) {
        this.residentialState = value;
    }

    /**
     * Gets the value of the secondName property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getSecondName() {
        return secondName;
    }

    /**
     * Sets the value of the secondName property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setSecondName(String value) {
        this.secondName = value;
    }

    /**
     * Gets the value of the siteCode property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getSiteCode() {
        return siteCode;
    }

    /**
     * Sets the value of the siteCode property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setSiteCode(String value) {
        this.siteCode = value;
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
     * Gets the value of the socialSecurityNoType property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getSocialSecurityNoType() {
        return socialSecurityNoType;
    }

    /**
     * Sets the value of the socialSecurityNoType property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setSocialSecurityNoType(String value) {
        this.socialSecurityNoType = value;
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
     * Gets the value of the thirdName property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getThirdName() {
        return thirdName;
    }

    /**
     * Sets the value of the thirdName property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setThirdName(String value) {
        this.thirdName = value;
    }

    /**
     * Gets the value of the title property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getTitle() {
        return title;
    }

    /**
     * Sets the value of the title property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setTitle(String value) {
        this.title = value;
    }

    /**
     * Gets the value of the workOrderPrefix property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getWorkOrderPrefix() {
        return workOrderPrefix;
    }

    /**
     * Sets the value of the workOrderPrefix property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setWorkOrderPrefix(String value) {
        this.workOrderPrefix = value;
    }

}
