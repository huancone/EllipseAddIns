
package com.mincom.enterpriseservice.ellipse.employee;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.AbstractDTO;


/**
 * <p>Java class for EmployeeServiceRetrieveForExtractRequestDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="EmployeeServiceRetrieveForExtractRequestDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractDTO">
 *       &lt;sequence>
 *         &lt;element name="coreEmployeeInd" type="{http://employee.ellipse.enterpriseservice.mincom.com}coreEmployeeInd" minOccurs="0"/>
 *         &lt;element name="employee" type="{http://employee.ellipse.enterpriseservice.mincom.com}employee" minOccurs="0"/>
 *         &lt;element name="firstName" type="{http://employee.ellipse.enterpriseservice.mincom.com}firstName" minOccurs="0"/>
 *         &lt;element name="includeDeceasedEmps" type="{http://employee.ellipse.enterpriseservice.mincom.com}includeDeceasedEmps" minOccurs="0"/>
 *         &lt;element name="includeTerminatedEmps" type="{http://employee.ellipse.enterpriseservice.mincom.com}includeTerminatedEmps" minOccurs="0"/>
 *         &lt;element name="lastName" type="{http://employee.ellipse.enterpriseservice.mincom.com}lastName" minOccurs="0"/>
 *         &lt;element name="lastRunDateTime" type="{http://employee.ellipse.enterpriseservice.mincom.com}lastRunDateTime" minOccurs="0"/>
 *         &lt;element name="paygroup" type="{http://employee.ellipse.enterpriseservice.mincom.com}paygroup" minOccurs="0"/>
 *         &lt;element name="payrollEmployeeInd" type="{http://employee.ellipse.enterpriseservice.mincom.com}payrollEmployeeInd" minOccurs="0"/>
 *         &lt;element name="personnelEmployeeInd" type="{http://employee.ellipse.enterpriseservice.mincom.com}personnelEmployeeInd" minOccurs="0"/>
 *         &lt;element name="personnelGroup" type="{http://employee.ellipse.enterpriseservice.mincom.com}personnelGroup" minOccurs="0"/>
 *         &lt;element name="personnelStatus" type="{http://employee.ellipse.enterpriseservice.mincom.com}personnelStatus" minOccurs="0"/>
 *         &lt;element name="physicalLocation" type="{http://employee.ellipse.enterpriseservice.mincom.com}physicalLocation" minOccurs="0"/>
 *         &lt;element name="position" type="{http://employee.ellipse.enterpriseservice.mincom.com}position" minOccurs="0"/>
 *         &lt;element name="preferredName" type="{http://employee.ellipse.enterpriseservice.mincom.com}preferredName" minOccurs="0"/>
 *         &lt;element name="resourceClass" type="{http://employee.ellipse.enterpriseservice.mincom.com}resourceClass" minOccurs="0"/>
 *         &lt;element name="resourceCode" type="{http://employee.ellipse.enterpriseservice.mincom.com}resourceCode" minOccurs="0"/>
 *         &lt;element name="socialInsuranceNumber" type="{http://employee.ellipse.enterpriseservice.mincom.com}socialInsuranceNumber" minOccurs="0"/>
 *         &lt;element name="socialSecurityNumber" type="{http://employee.ellipse.enterpriseservice.mincom.com}socialSecurityNumber" minOccurs="0"/>
 *         &lt;element name="workGroup" type="{http://employee.ellipse.enterpriseservice.mincom.com}workGroup" minOccurs="0"/>
 *         &lt;element name="workGroupCrew" type="{http://employee.ellipse.enterpriseservice.mincom.com}workGroupCrew" minOccurs="0"/>
 *         &lt;element name="workGroupStartDate" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *         &lt;element name="workGroupStopDate" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *         &lt;element name="workLocation" type="{http://employee.ellipse.enterpriseservice.mincom.com}workLocation" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/extension>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "EmployeeServiceRetrieveForExtractRequestDTO", propOrder = {
    "coreEmployeeInd",
    "employee",
    "firstName",
    "includeDeceasedEmps",
    "includeTerminatedEmps",
    "lastName",
    "lastRunDateTime",
    "paygroup",
    "payrollEmployeeInd",
    "personnelEmployeeInd",
    "personnelGroup",
    "personnelStatus",
    "physicalLocation",
    "position",
    "preferredName",
    "resourceClass",
    "resourceCode",
    "socialInsuranceNumber",
    "socialSecurityNumber",
    "workGroup",
    "workGroupCrew",
    "workGroupStartDate",
    "workGroupStopDate",
    "workLocation"
})
public class EmployeeServiceRetrieveForExtractRequestDTO
    extends AbstractDTO
{

    protected Boolean coreEmployeeInd;
    protected String employee;
    protected String firstName;
    protected Boolean includeDeceasedEmps;
    protected Boolean includeTerminatedEmps;
    protected String lastName;
    protected String lastRunDateTime;
    protected String paygroup;
    protected Boolean payrollEmployeeInd;
    protected Boolean personnelEmployeeInd;
    protected String personnelGroup;
    protected String personnelStatus;
    protected String physicalLocation;
    protected String position;
    protected String preferredName;
    protected String resourceClass;
    protected String resourceCode;
    protected String socialInsuranceNumber;
    protected String socialSecurityNumber;
    protected String workGroup;
    protected String workGroupCrew;
    protected String workGroupStartDate;
    protected String workGroupStopDate;
    protected String workLocation;

    /**
     * Gets the value of the coreEmployeeInd property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isCoreEmployeeInd() {
        return coreEmployeeInd;
    }

    /**
     * Sets the value of the coreEmployeeInd property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setCoreEmployeeInd(Boolean value) {
        this.coreEmployeeInd = value;
    }

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
     * Gets the value of the includeDeceasedEmps property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isIncludeDeceasedEmps() {
        return includeDeceasedEmps;
    }

    /**
     * Sets the value of the includeDeceasedEmps property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setIncludeDeceasedEmps(Boolean value) {
        this.includeDeceasedEmps = value;
    }

    /**
     * Gets the value of the includeTerminatedEmps property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isIncludeTerminatedEmps() {
        return includeTerminatedEmps;
    }

    /**
     * Sets the value of the includeTerminatedEmps property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setIncludeTerminatedEmps(Boolean value) {
        this.includeTerminatedEmps = value;
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
     * Gets the value of the lastRunDateTime property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getLastRunDateTime() {
        return lastRunDateTime;
    }

    /**
     * Sets the value of the lastRunDateTime property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setLastRunDateTime(String value) {
        this.lastRunDateTime = value;
    }

    /**
     * Gets the value of the paygroup property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPaygroup() {
        return paygroup;
    }

    /**
     * Sets the value of the paygroup property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPaygroup(String value) {
        this.paygroup = value;
    }

    /**
     * Gets the value of the payrollEmployeeInd property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isPayrollEmployeeInd() {
        return payrollEmployeeInd;
    }

    /**
     * Sets the value of the payrollEmployeeInd property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setPayrollEmployeeInd(Boolean value) {
        this.payrollEmployeeInd = value;
    }

    /**
     * Gets the value of the personnelEmployeeInd property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isPersonnelEmployeeInd() {
        return personnelEmployeeInd;
    }

    /**
     * Sets the value of the personnelEmployeeInd property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setPersonnelEmployeeInd(Boolean value) {
        this.personnelEmployeeInd = value;
    }

    /**
     * Gets the value of the personnelGroup property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPersonnelGroup() {
        return personnelGroup;
    }

    /**
     * Sets the value of the personnelGroup property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPersonnelGroup(String value) {
        this.personnelGroup = value;
    }

    /**
     * Gets the value of the personnelStatus property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPersonnelStatus() {
        return personnelStatus;
    }

    /**
     * Sets the value of the personnelStatus property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPersonnelStatus(String value) {
        this.personnelStatus = value;
    }

    /**
     * Gets the value of the physicalLocation property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPhysicalLocation() {
        return physicalLocation;
    }

    /**
     * Sets the value of the physicalLocation property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPhysicalLocation(String value) {
        this.physicalLocation = value;
    }

    /**
     * Gets the value of the position property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPosition() {
        return position;
    }

    /**
     * Sets the value of the position property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPosition(String value) {
        this.position = value;
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
     * Gets the value of the resourceClass property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getResourceClass() {
        return resourceClass;
    }

    /**
     * Sets the value of the resourceClass property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setResourceClass(String value) {
        this.resourceClass = value;
    }

    /**
     * Gets the value of the resourceCode property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getResourceCode() {
        return resourceCode;
    }

    /**
     * Sets the value of the resourceCode property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setResourceCode(String value) {
        this.resourceCode = value;
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
     * Gets the value of the workGroup property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getWorkGroup() {
        return workGroup;
    }

    /**
     * Sets the value of the workGroup property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setWorkGroup(String value) {
        this.workGroup = value;
    }

    /**
     * Gets the value of the workGroupCrew property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getWorkGroupCrew() {
        return workGroupCrew;
    }

    /**
     * Sets the value of the workGroupCrew property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setWorkGroupCrew(String value) {
        this.workGroupCrew = value;
    }

    /**
     * Gets the value of the workGroupStartDate property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getWorkGroupStartDate() {
        return workGroupStartDate;
    }

    /**
     * Sets the value of the workGroupStartDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setWorkGroupStartDate(String value) {
        this.workGroupStartDate = value;
    }

    /**
     * Gets the value of the workGroupStopDate property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getWorkGroupStopDate() {
        return workGroupStopDate;
    }

    /**
     * Sets the value of the workGroupStopDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setWorkGroupStopDate(String value) {
        this.workGroupStopDate = value;
    }

    /**
     * Gets the value of the workLocation property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getWorkLocation() {
        return workLocation;
    }

    /**
     * Sets the value of the workLocation property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setWorkLocation(String value) {
        this.workLocation = value;
    }

}
