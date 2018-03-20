
package com.mincom.enterpriseservice.ellipse.employee;

import java.math.BigDecimal;
import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.AbstractReplyDTO;


/**
 * <p>Java class for EmployeeServiceFetchLeaveHeaderReplyDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="EmployeeServiceFetchLeaveHeaderReplyDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractReplyDTO">
 *       &lt;sequence>
 *         &lt;element name="awardCode" type="{http://employee.ellipse.enterpriseservice.mincom.com}awardCode" minOccurs="0"/>
 *         &lt;element name="awardCodeDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}awardCodeDesc" minOccurs="0"/>
 *         &lt;element name="contractHours" type="{http://employee.ellipse.enterpriseservice.mincom.com}contractHours" minOccurs="0"/>
 *         &lt;element name="contractMinutes" type="{http://employee.ellipse.enterpriseservice.mincom.com}contractMinutes" minOccurs="0"/>
 *         &lt;element name="employee" type="{http://employee.ellipse.enterpriseservice.mincom.com}employee" minOccurs="0"/>
 *         &lt;element name="employeeClass" type="{http://employee.ellipse.enterpriseservice.mincom.com}employeeClass" minOccurs="0"/>
 *         &lt;element name="employeeClassDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}employeeClassDesc" minOccurs="0"/>
 *         &lt;element name="employeeFormattedName" type="{http://employee.ellipse.enterpriseservice.mincom.com}employeeFormattedName" minOccurs="0"/>
 *         &lt;element name="employeeType" type="{http://employee.ellipse.enterpriseservice.mincom.com}employeeType" minOccurs="0"/>
 *         &lt;element name="employeeTypeDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}employeeTypeDesc" minOccurs="0"/>
 *         &lt;element name="essUserInd" type="{http://employee.ellipse.enterpriseservice.mincom.com}essUserInd" minOccurs="0"/>
 *         &lt;element name="firstName" type="{http://employee.ellipse.enterpriseservice.mincom.com}firstName" minOccurs="0"/>
 *         &lt;element name="hireDate" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *         &lt;element name="lastName" type="{http://employee.ellipse.enterpriseservice.mincom.com}lastName" minOccurs="0"/>
 *         &lt;element name="leaveForecastDate" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *         &lt;element name="persEmpStatus" type="{http://employee.ellipse.enterpriseservice.mincom.com}persEmpStatus" minOccurs="0"/>
 *         &lt;element name="persEmpStatusDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}persEmpStatusDesc" minOccurs="0"/>
 *         &lt;element name="personnelGroup" type="{http://employee.ellipse.enterpriseservice.mincom.com}personnelGroup" minOccurs="0"/>
 *         &lt;element name="personnelGroupDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}personnelGroupDesc" minOccurs="0"/>
 *         &lt;element name="personnelStatus" type="{http://employee.ellipse.enterpriseservice.mincom.com}personnelStatus" minOccurs="0"/>
 *         &lt;element name="personnelStatusDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}personnelStatusDesc" minOccurs="0"/>
 *         &lt;element name="physicalLocation" type="{http://employee.ellipse.enterpriseservice.mincom.com}physicalLocation" minOccurs="0"/>
 *         &lt;element name="physicalLocationDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}physicalLocationDesc" minOccurs="0"/>
 *         &lt;element name="position" type="{http://employee.ellipse.enterpriseservice.mincom.com}position" minOccurs="0"/>
 *         &lt;element name="positionDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}positionDesc" minOccurs="0"/>
 *         &lt;element name="primRepCode" type="{http://employee.ellipse.enterpriseservice.mincom.com}primRepCode" minOccurs="0"/>
 *         &lt;element name="primRepCodeDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}primRepCodeDesc" minOccurs="0"/>
 *         &lt;element name="professionalServiceDate" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *         &lt;element name="secondName" type="{http://employee.ellipse.enterpriseservice.mincom.com}secondName" minOccurs="0"/>
 *         &lt;element name="serviceDate" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *         &lt;element name="thirdName" type="{http://employee.ellipse.enterpriseservice.mincom.com}thirdName" minOccurs="0"/>
 *         &lt;element name="title" type="{http://employee.ellipse.enterpriseservice.mincom.com}title" minOccurs="0"/>
 *         &lt;element name="titleDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}titleDesc" minOccurs="0"/>
 *         &lt;element name="workGroup" type="{http://employee.ellipse.enterpriseservice.mincom.com}workGroup" minOccurs="0"/>
 *         &lt;element name="workGroupCrew" type="{http://employee.ellipse.enterpriseservice.mincom.com}workGroupCrew" minOccurs="0"/>
 *         &lt;element name="workGroupCrewDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}workGroupCrewDesc" minOccurs="0"/>
 *         &lt;element name="workGroupDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}workGroupDesc" minOccurs="0"/>
 *         &lt;element name="workLocation" type="{http://employee.ellipse.enterpriseservice.mincom.com}workLocation" minOccurs="0"/>
 *         &lt;element name="workLocationDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}workLocationDesc" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/extension>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "EmployeeServiceFetchLeaveHeaderReplyDTO", propOrder = {
    "awardCode",
    "awardCodeDesc",
    "contractHours",
    "contractMinutes",
    "employee",
    "employeeClass",
    "employeeClassDesc",
    "employeeFormattedName",
    "employeeType",
    "employeeTypeDesc",
    "essUserInd",
    "firstName",
    "hireDate",
    "lastName",
    "leaveForecastDate",
    "persEmpStatus",
    "persEmpStatusDesc",
    "personnelGroup",
    "personnelGroupDesc",
    "personnelStatus",
    "personnelStatusDesc",
    "physicalLocation",
    "physicalLocationDesc",
    "position",
    "positionDesc",
    "primRepCode",
    "primRepCodeDesc",
    "professionalServiceDate",
    "secondName",
    "serviceDate",
    "thirdName",
    "title",
    "titleDesc",
    "workGroup",
    "workGroupCrew",
    "workGroupCrewDesc",
    "workGroupDesc",
    "workLocation",
    "workLocationDesc"
})
public class EmployeeServiceFetchLeaveHeaderReplyDTO
    extends AbstractReplyDTO
{

    protected String awardCode;
    protected String awardCodeDesc;
    protected BigDecimal contractHours;
    protected BigDecimal contractMinutes;
    protected String employee;
    protected String employeeClass;
    protected String employeeClassDesc;
    protected String employeeFormattedName;
    protected String employeeType;
    protected String employeeTypeDesc;
    protected Boolean essUserInd;
    protected String firstName;
    protected String hireDate;
    protected String lastName;
    protected String leaveForecastDate;
    protected String persEmpStatus;
    protected String persEmpStatusDesc;
    protected String personnelGroup;
    protected String personnelGroupDesc;
    protected String personnelStatus;
    protected String personnelStatusDesc;
    protected String physicalLocation;
    protected String physicalLocationDesc;
    protected String position;
    protected String positionDesc;
    protected String primRepCode;
    protected String primRepCodeDesc;
    protected String professionalServiceDate;
    protected String secondName;
    protected String serviceDate;
    protected String thirdName;
    protected String title;
    protected String titleDesc;
    protected String workGroup;
    protected String workGroupCrew;
    protected String workGroupCrewDesc;
    protected String workGroupDesc;
    protected String workLocation;
    protected String workLocationDesc;

    /**
     * Gets the value of the awardCode property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getAwardCode() {
        return awardCode;
    }

    /**
     * Sets the value of the awardCode property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setAwardCode(String value) {
        this.awardCode = value;
    }

    /**
     * Gets the value of the awardCodeDesc property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getAwardCodeDesc() {
        return awardCodeDesc;
    }

    /**
     * Sets the value of the awardCodeDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setAwardCodeDesc(String value) {
        this.awardCodeDesc = value;
    }

    /**
     * Gets the value of the contractHours property.
     * 
     * @return
     *     possible object is
     *     {@link BigDecimal }
     *     
     */
    public BigDecimal getContractHours() {
        return contractHours;
    }

    /**
     * Sets the value of the contractHours property.
     * 
     * @param value
     *     allowed object is
     *     {@link BigDecimal }
     *     
     */
    public void setContractHours(BigDecimal value) {
        this.contractHours = value;
    }

    /**
     * Gets the value of the contractMinutes property.
     * 
     * @return
     *     possible object is
     *     {@link BigDecimal }
     *     
     */
    public BigDecimal getContractMinutes() {
        return contractMinutes;
    }

    /**
     * Sets the value of the contractMinutes property.
     * 
     * @param value
     *     allowed object is
     *     {@link BigDecimal }
     *     
     */
    public void setContractMinutes(BigDecimal value) {
        this.contractMinutes = value;
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
     * Gets the value of the employeeClass property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getEmployeeClass() {
        return employeeClass;
    }

    /**
     * Sets the value of the employeeClass property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setEmployeeClass(String value) {
        this.employeeClass = value;
    }

    /**
     * Gets the value of the employeeClassDesc property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getEmployeeClassDesc() {
        return employeeClassDesc;
    }

    /**
     * Sets the value of the employeeClassDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setEmployeeClassDesc(String value) {
        this.employeeClassDesc = value;
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
     * Gets the value of the employeeType property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getEmployeeType() {
        return employeeType;
    }

    /**
     * Sets the value of the employeeType property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setEmployeeType(String value) {
        this.employeeType = value;
    }

    /**
     * Gets the value of the employeeTypeDesc property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getEmployeeTypeDesc() {
        return employeeTypeDesc;
    }

    /**
     * Sets the value of the employeeTypeDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setEmployeeTypeDesc(String value) {
        this.employeeTypeDesc = value;
    }

    /**
     * Gets the value of the essUserInd property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isEssUserInd() {
        return essUserInd;
    }

    /**
     * Sets the value of the essUserInd property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setEssUserInd(Boolean value) {
        this.essUserInd = value;
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
     * Gets the value of the hireDate property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getHireDate() {
        return hireDate;
    }

    /**
     * Sets the value of the hireDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setHireDate(String value) {
        this.hireDate = value;
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
     * Gets the value of the leaveForecastDate property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getLeaveForecastDate() {
        return leaveForecastDate;
    }

    /**
     * Sets the value of the leaveForecastDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setLeaveForecastDate(String value) {
        this.leaveForecastDate = value;
    }

    /**
     * Gets the value of the persEmpStatus property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPersEmpStatus() {
        return persEmpStatus;
    }

    /**
     * Sets the value of the persEmpStatus property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPersEmpStatus(String value) {
        this.persEmpStatus = value;
    }

    /**
     * Gets the value of the persEmpStatusDesc property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPersEmpStatusDesc() {
        return persEmpStatusDesc;
    }

    /**
     * Sets the value of the persEmpStatusDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPersEmpStatusDesc(String value) {
        this.persEmpStatusDesc = value;
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
     * Gets the value of the personnelGroupDesc property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPersonnelGroupDesc() {
        return personnelGroupDesc;
    }

    /**
     * Sets the value of the personnelGroupDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPersonnelGroupDesc(String value) {
        this.personnelGroupDesc = value;
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
     * Gets the value of the personnelStatusDesc property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPersonnelStatusDesc() {
        return personnelStatusDesc;
    }

    /**
     * Sets the value of the personnelStatusDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPersonnelStatusDesc(String value) {
        this.personnelStatusDesc = value;
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
     * Gets the value of the physicalLocationDesc property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPhysicalLocationDesc() {
        return physicalLocationDesc;
    }

    /**
     * Sets the value of the physicalLocationDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPhysicalLocationDesc(String value) {
        this.physicalLocationDesc = value;
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
     * Gets the value of the positionDesc property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPositionDesc() {
        return positionDesc;
    }

    /**
     * Sets the value of the positionDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPositionDesc(String value) {
        this.positionDesc = value;
    }

    /**
     * Gets the value of the primRepCode property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPrimRepCode() {
        return primRepCode;
    }

    /**
     * Sets the value of the primRepCode property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPrimRepCode(String value) {
        this.primRepCode = value;
    }

    /**
     * Gets the value of the primRepCodeDesc property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPrimRepCodeDesc() {
        return primRepCodeDesc;
    }

    /**
     * Sets the value of the primRepCodeDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPrimRepCodeDesc(String value) {
        this.primRepCodeDesc = value;
    }

    /**
     * Gets the value of the professionalServiceDate property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getProfessionalServiceDate() {
        return professionalServiceDate;
    }

    /**
     * Sets the value of the professionalServiceDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setProfessionalServiceDate(String value) {
        this.professionalServiceDate = value;
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
     * Gets the value of the serviceDate property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getServiceDate() {
        return serviceDate;
    }

    /**
     * Sets the value of the serviceDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setServiceDate(String value) {
        this.serviceDate = value;
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
     * Gets the value of the workGroupCrewDesc property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getWorkGroupCrewDesc() {
        return workGroupCrewDesc;
    }

    /**
     * Sets the value of the workGroupCrewDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setWorkGroupCrewDesc(String value) {
        this.workGroupCrewDesc = value;
    }

    /**
     * Gets the value of the workGroupDesc property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getWorkGroupDesc() {
        return workGroupDesc;
    }

    /**
     * Sets the value of the workGroupDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setWorkGroupDesc(String value) {
        this.workGroupDesc = value;
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

    /**
     * Gets the value of the workLocationDesc property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getWorkLocationDesc() {
        return workLocationDesc;
    }

    /**
     * Sets the value of the workLocationDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setWorkLocationDesc(String value) {
        this.workLocationDesc = value;
    }

}
