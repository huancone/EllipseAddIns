
package com.mincom.enterpriseservice.ellipse.employee;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.AbstractRequiredAttributesDTO;


/**
 * <p>Java class for EmployeeServiceFetchLeaveHeaderRequiredAttributesDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="EmployeeServiceFetchLeaveHeaderRequiredAttributesDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractRequiredAttributesDTO">
 *       &lt;sequence>
 *         &lt;element name="returnAwardCode" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnAwardCodeDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnContractHours" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnContractMinutes" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnEmployee" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnEmployeeClass" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnEmployeeClassDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnEmployeeFormattedName" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnEmployeeType" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnEmployeeTypeDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnEssUserInd" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnFirstName" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnHireDate" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnLastName" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnLeaveForecastDate" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersEmpStatus" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersEmpStatusDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelGroup" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelGroupDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelStatus" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelStatusDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPhysicalLocation" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPhysicalLocationDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPosition" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPositionDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPrimRepCode" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPrimRepCodeDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnProfessionalServiceDate" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSecondName" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnServiceDate" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnThirdName" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnTitle" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnTitleDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnWorkGroup" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnWorkGroupCrew" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnWorkGroupCrewDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnWorkGroupDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnWorkLocation" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnWorkLocationDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/extension>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "EmployeeServiceFetchLeaveHeaderRequiredAttributesDTO", propOrder = {
    "returnAwardCode",
    "returnAwardCodeDesc",
    "returnContractHours",
    "returnContractMinutes",
    "returnEmployee",
    "returnEmployeeClass",
    "returnEmployeeClassDesc",
    "returnEmployeeFormattedName",
    "returnEmployeeType",
    "returnEmployeeTypeDesc",
    "returnEssUserInd",
    "returnFirstName",
    "returnHireDate",
    "returnLastName",
    "returnLeaveForecastDate",
    "returnPersEmpStatus",
    "returnPersEmpStatusDesc",
    "returnPersonnelGroup",
    "returnPersonnelGroupDesc",
    "returnPersonnelStatus",
    "returnPersonnelStatusDesc",
    "returnPhysicalLocation",
    "returnPhysicalLocationDesc",
    "returnPosition",
    "returnPositionDesc",
    "returnPrimRepCode",
    "returnPrimRepCodeDesc",
    "returnProfessionalServiceDate",
    "returnSecondName",
    "returnServiceDate",
    "returnThirdName",
    "returnTitle",
    "returnTitleDesc",
    "returnWorkGroup",
    "returnWorkGroupCrew",
    "returnWorkGroupCrewDesc",
    "returnWorkGroupDesc",
    "returnWorkLocation",
    "returnWorkLocationDesc"
})
public class EmployeeServiceFetchLeaveHeaderRequiredAttributesDTO
    extends AbstractRequiredAttributesDTO
{

    protected Boolean returnAwardCode;
    protected Boolean returnAwardCodeDesc;
    protected Boolean returnContractHours;
    protected Boolean returnContractMinutes;
    protected Boolean returnEmployee;
    protected Boolean returnEmployeeClass;
    protected Boolean returnEmployeeClassDesc;
    protected Boolean returnEmployeeFormattedName;
    protected Boolean returnEmployeeType;
    protected Boolean returnEmployeeTypeDesc;
    protected Boolean returnEssUserInd;
    protected Boolean returnFirstName;
    protected Boolean returnHireDate;
    protected Boolean returnLastName;
    protected Boolean returnLeaveForecastDate;
    protected Boolean returnPersEmpStatus;
    protected Boolean returnPersEmpStatusDesc;
    protected Boolean returnPersonnelGroup;
    protected Boolean returnPersonnelGroupDesc;
    protected Boolean returnPersonnelStatus;
    protected Boolean returnPersonnelStatusDesc;
    protected Boolean returnPhysicalLocation;
    protected Boolean returnPhysicalLocationDesc;
    protected Boolean returnPosition;
    protected Boolean returnPositionDesc;
    protected Boolean returnPrimRepCode;
    protected Boolean returnPrimRepCodeDesc;
    protected Boolean returnProfessionalServiceDate;
    protected Boolean returnSecondName;
    protected Boolean returnServiceDate;
    protected Boolean returnThirdName;
    protected Boolean returnTitle;
    protected Boolean returnTitleDesc;
    protected Boolean returnWorkGroup;
    protected Boolean returnWorkGroupCrew;
    protected Boolean returnWorkGroupCrewDesc;
    protected Boolean returnWorkGroupDesc;
    protected Boolean returnWorkLocation;
    protected Boolean returnWorkLocationDesc;

    /**
     * Gets the value of the returnAwardCode property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnAwardCode() {
        return returnAwardCode;
    }

    /**
     * Sets the value of the returnAwardCode property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnAwardCode(Boolean value) {
        this.returnAwardCode = value;
    }

    /**
     * Gets the value of the returnAwardCodeDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnAwardCodeDesc() {
        return returnAwardCodeDesc;
    }

    /**
     * Sets the value of the returnAwardCodeDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnAwardCodeDesc(Boolean value) {
        this.returnAwardCodeDesc = value;
    }

    /**
     * Gets the value of the returnContractHours property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnContractHours() {
        return returnContractHours;
    }

    /**
     * Sets the value of the returnContractHours property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnContractHours(Boolean value) {
        this.returnContractHours = value;
    }

    /**
     * Gets the value of the returnContractMinutes property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnContractMinutes() {
        return returnContractMinutes;
    }

    /**
     * Sets the value of the returnContractMinutes property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnContractMinutes(Boolean value) {
        this.returnContractMinutes = value;
    }

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
     * Gets the value of the returnEmployeeClass property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnEmployeeClass() {
        return returnEmployeeClass;
    }

    /**
     * Sets the value of the returnEmployeeClass property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnEmployeeClass(Boolean value) {
        this.returnEmployeeClass = value;
    }

    /**
     * Gets the value of the returnEmployeeClassDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnEmployeeClassDesc() {
        return returnEmployeeClassDesc;
    }

    /**
     * Sets the value of the returnEmployeeClassDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnEmployeeClassDesc(Boolean value) {
        this.returnEmployeeClassDesc = value;
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
     * Gets the value of the returnEmployeeType property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnEmployeeType() {
        return returnEmployeeType;
    }

    /**
     * Sets the value of the returnEmployeeType property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnEmployeeType(Boolean value) {
        this.returnEmployeeType = value;
    }

    /**
     * Gets the value of the returnEmployeeTypeDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnEmployeeTypeDesc() {
        return returnEmployeeTypeDesc;
    }

    /**
     * Sets the value of the returnEmployeeTypeDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnEmployeeTypeDesc(Boolean value) {
        this.returnEmployeeTypeDesc = value;
    }

    /**
     * Gets the value of the returnEssUserInd property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnEssUserInd() {
        return returnEssUserInd;
    }

    /**
     * Sets the value of the returnEssUserInd property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnEssUserInd(Boolean value) {
        this.returnEssUserInd = value;
    }

    /**
     * Gets the value of the returnFirstName property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnFirstName() {
        return returnFirstName;
    }

    /**
     * Sets the value of the returnFirstName property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnFirstName(Boolean value) {
        this.returnFirstName = value;
    }

    /**
     * Gets the value of the returnHireDate property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnHireDate() {
        return returnHireDate;
    }

    /**
     * Sets the value of the returnHireDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnHireDate(Boolean value) {
        this.returnHireDate = value;
    }

    /**
     * Gets the value of the returnLastName property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnLastName() {
        return returnLastName;
    }

    /**
     * Sets the value of the returnLastName property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnLastName(Boolean value) {
        this.returnLastName = value;
    }

    /**
     * Gets the value of the returnLeaveForecastDate property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnLeaveForecastDate() {
        return returnLeaveForecastDate;
    }

    /**
     * Sets the value of the returnLeaveForecastDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnLeaveForecastDate(Boolean value) {
        this.returnLeaveForecastDate = value;
    }

    /**
     * Gets the value of the returnPersEmpStatus property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersEmpStatus() {
        return returnPersEmpStatus;
    }

    /**
     * Sets the value of the returnPersEmpStatus property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersEmpStatus(Boolean value) {
        this.returnPersEmpStatus = value;
    }

    /**
     * Gets the value of the returnPersEmpStatusDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersEmpStatusDesc() {
        return returnPersEmpStatusDesc;
    }

    /**
     * Sets the value of the returnPersEmpStatusDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersEmpStatusDesc(Boolean value) {
        this.returnPersEmpStatusDesc = value;
    }

    /**
     * Gets the value of the returnPersonnelGroup property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelGroup() {
        return returnPersonnelGroup;
    }

    /**
     * Sets the value of the returnPersonnelGroup property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelGroup(Boolean value) {
        this.returnPersonnelGroup = value;
    }

    /**
     * Gets the value of the returnPersonnelGroupDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelGroupDesc() {
        return returnPersonnelGroupDesc;
    }

    /**
     * Sets the value of the returnPersonnelGroupDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelGroupDesc(Boolean value) {
        this.returnPersonnelGroupDesc = value;
    }

    /**
     * Gets the value of the returnPersonnelStatus property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelStatus() {
        return returnPersonnelStatus;
    }

    /**
     * Sets the value of the returnPersonnelStatus property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelStatus(Boolean value) {
        this.returnPersonnelStatus = value;
    }

    /**
     * Gets the value of the returnPersonnelStatusDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelStatusDesc() {
        return returnPersonnelStatusDesc;
    }

    /**
     * Sets the value of the returnPersonnelStatusDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelStatusDesc(Boolean value) {
        this.returnPersonnelStatusDesc = value;
    }

    /**
     * Gets the value of the returnPhysicalLocation property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPhysicalLocation() {
        return returnPhysicalLocation;
    }

    /**
     * Sets the value of the returnPhysicalLocation property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPhysicalLocation(Boolean value) {
        this.returnPhysicalLocation = value;
    }

    /**
     * Gets the value of the returnPhysicalLocationDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPhysicalLocationDesc() {
        return returnPhysicalLocationDesc;
    }

    /**
     * Sets the value of the returnPhysicalLocationDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPhysicalLocationDesc(Boolean value) {
        this.returnPhysicalLocationDesc = value;
    }

    /**
     * Gets the value of the returnPosition property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPosition() {
        return returnPosition;
    }

    /**
     * Sets the value of the returnPosition property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPosition(Boolean value) {
        this.returnPosition = value;
    }

    /**
     * Gets the value of the returnPositionDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPositionDesc() {
        return returnPositionDesc;
    }

    /**
     * Sets the value of the returnPositionDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPositionDesc(Boolean value) {
        this.returnPositionDesc = value;
    }

    /**
     * Gets the value of the returnPrimRepCode property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPrimRepCode() {
        return returnPrimRepCode;
    }

    /**
     * Sets the value of the returnPrimRepCode property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPrimRepCode(Boolean value) {
        this.returnPrimRepCode = value;
    }

    /**
     * Gets the value of the returnPrimRepCodeDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPrimRepCodeDesc() {
        return returnPrimRepCodeDesc;
    }

    /**
     * Sets the value of the returnPrimRepCodeDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPrimRepCodeDesc(Boolean value) {
        this.returnPrimRepCodeDesc = value;
    }

    /**
     * Gets the value of the returnProfessionalServiceDate property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnProfessionalServiceDate() {
        return returnProfessionalServiceDate;
    }

    /**
     * Sets the value of the returnProfessionalServiceDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnProfessionalServiceDate(Boolean value) {
        this.returnProfessionalServiceDate = value;
    }

    /**
     * Gets the value of the returnSecondName property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSecondName() {
        return returnSecondName;
    }

    /**
     * Sets the value of the returnSecondName property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSecondName(Boolean value) {
        this.returnSecondName = value;
    }

    /**
     * Gets the value of the returnServiceDate property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnServiceDate() {
        return returnServiceDate;
    }

    /**
     * Sets the value of the returnServiceDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnServiceDate(Boolean value) {
        this.returnServiceDate = value;
    }

    /**
     * Gets the value of the returnThirdName property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnThirdName() {
        return returnThirdName;
    }

    /**
     * Sets the value of the returnThirdName property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnThirdName(Boolean value) {
        this.returnThirdName = value;
    }

    /**
     * Gets the value of the returnTitle property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnTitle() {
        return returnTitle;
    }

    /**
     * Sets the value of the returnTitle property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnTitle(Boolean value) {
        this.returnTitle = value;
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
     * Gets the value of the returnWorkGroup property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnWorkGroup() {
        return returnWorkGroup;
    }

    /**
     * Sets the value of the returnWorkGroup property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnWorkGroup(Boolean value) {
        this.returnWorkGroup = value;
    }

    /**
     * Gets the value of the returnWorkGroupCrew property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnWorkGroupCrew() {
        return returnWorkGroupCrew;
    }

    /**
     * Sets the value of the returnWorkGroupCrew property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnWorkGroupCrew(Boolean value) {
        this.returnWorkGroupCrew = value;
    }

    /**
     * Gets the value of the returnWorkGroupCrewDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnWorkGroupCrewDesc() {
        return returnWorkGroupCrewDesc;
    }

    /**
     * Sets the value of the returnWorkGroupCrewDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnWorkGroupCrewDesc(Boolean value) {
        this.returnWorkGroupCrewDesc = value;
    }

    /**
     * Gets the value of the returnWorkGroupDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnWorkGroupDesc() {
        return returnWorkGroupDesc;
    }

    /**
     * Sets the value of the returnWorkGroupDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnWorkGroupDesc(Boolean value) {
        this.returnWorkGroupDesc = value;
    }

    /**
     * Gets the value of the returnWorkLocation property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnWorkLocation() {
        return returnWorkLocation;
    }

    /**
     * Sets the value of the returnWorkLocation property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnWorkLocation(Boolean value) {
        this.returnWorkLocation = value;
    }

    /**
     * Gets the value of the returnWorkLocationDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnWorkLocationDesc() {
        return returnWorkLocationDesc;
    }

    /**
     * Sets the value of the returnWorkLocationDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnWorkLocationDesc(Boolean value) {
        this.returnWorkLocationDesc = value;
    }

}
