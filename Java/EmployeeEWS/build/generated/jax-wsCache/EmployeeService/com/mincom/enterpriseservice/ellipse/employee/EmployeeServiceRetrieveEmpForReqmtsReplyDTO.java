
package com.mincom.enterpriseservice.ellipse.employee;

import java.math.BigDecimal;
import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.AbstractReplyDTO;


/**
 * <p>Java class for EmployeeServiceRetrieveEmpForReqmtsReplyDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="EmployeeServiceRetrieveEmpForReqmtsReplyDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractReplyDTO">
 *       &lt;sequence>
 *         &lt;element name="activity" type="{http://employee.ellipse.enterpriseservice.mincom.com}activity" minOccurs="0"/>
 *         &lt;element name="activityDescription" type="{http://employee.ellipse.enterpriseservice.mincom.com}activityDescription" minOccurs="0"/>
 *         &lt;element name="courseMatch" type="{http://employee.ellipse.enterpriseservice.mincom.com}ArrayOfString" minOccurs="0"/>
 *         &lt;element name="courseMatchedWeightedPercentage" type="{http://employee.ellipse.enterpriseservice.mincom.com}courseMatchedWeightedPercentage" minOccurs="0"/>
 *         &lt;element name="courseNumberMatched" type="{http://employee.ellipse.enterpriseservice.mincom.com}courseNumberMatched" minOccurs="0"/>
 *         &lt;element name="courseTotalScore" type="{http://employee.ellipse.enterpriseservice.mincom.com}courseTotalScore" minOccurs="0"/>
 *         &lt;element name="employee" type="{http://employee.ellipse.enterpriseservice.mincom.com}employee" minOccurs="0"/>
 *         &lt;element name="employeeFormattedName" type="{http://employee.ellipse.enterpriseservice.mincom.com}employeeFormattedName" minOccurs="0"/>
 *         &lt;element name="jobClassLevel" type="{http://employee.ellipse.enterpriseservice.mincom.com}jobClassLevel" minOccurs="0"/>
 *         &lt;element name="jobClassLevelDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}jobClassLevelDesc" minOccurs="0"/>
 *         &lt;element name="matchedWeightedPercentage" type="{http://employee.ellipse.enterpriseservice.mincom.com}matchedWeightedPercentage" minOccurs="0"/>
 *         &lt;element name="numberMatched" type="{http://employee.ellipse.enterpriseservice.mincom.com}numberMatched" minOccurs="0"/>
 *         &lt;element name="payrollEmployeeInd" type="{http://employee.ellipse.enterpriseservice.mincom.com}payrollEmployeeInd" minOccurs="0"/>
 *         &lt;element name="persEmpStatusDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}persEmpStatusDesc" minOccurs="0"/>
 *         &lt;element name="personnelStatusDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}personnelStatusDesc" minOccurs="0"/>
 *         &lt;element name="physicalLocation" type="{http://employee.ellipse.enterpriseservice.mincom.com}physicalLocation" minOccurs="0"/>
 *         &lt;element name="physicalLocationDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}physicalLocationDesc" minOccurs="0"/>
 *         &lt;element name="position" type="{http://employee.ellipse.enterpriseservice.mincom.com}position" minOccurs="0"/>
 *         &lt;element name="positionDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}positionDesc" minOccurs="0"/>
 *         &lt;element name="positionMatch" type="{http://employee.ellipse.enterpriseservice.mincom.com}ArrayOfString" minOccurs="0"/>
 *         &lt;element name="positionMatchedWeightedPercentage" type="{http://employee.ellipse.enterpriseservice.mincom.com}positionMatchedWeightedPercentage" minOccurs="0"/>
 *         &lt;element name="positionNumberMatched" type="{http://employee.ellipse.enterpriseservice.mincom.com}positionNumberMatched" minOccurs="0"/>
 *         &lt;element name="positionTotalScore" type="{http://employee.ellipse.enterpriseservice.mincom.com}positionTotalScore" minOccurs="0"/>
 *         &lt;element name="resourceCompetency" type="{http://employee.ellipse.enterpriseservice.mincom.com}ArrayOfString" minOccurs="0"/>
 *         &lt;element name="resourceCompetencyDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}ArrayOfString" minOccurs="0"/>
 *         &lt;element name="resourceMatch" type="{http://employee.ellipse.enterpriseservice.mincom.com}ArrayOfString" minOccurs="0"/>
 *         &lt;element name="resourceMatchedWeightedPercentage" type="{http://employee.ellipse.enterpriseservice.mincom.com}resourceMatchedWeightedPercentage" minOccurs="0"/>
 *         &lt;element name="resourceNumberMatched" type="{http://employee.ellipse.enterpriseservice.mincom.com}resourceNumberMatched" minOccurs="0"/>
 *         &lt;element name="resourceTotalScore" type="{http://employee.ellipse.enterpriseservice.mincom.com}resourceTotalScore" minOccurs="0"/>
 *         &lt;element name="totalScore" type="{http://employee.ellipse.enterpriseservice.mincom.com}totalScore" minOccurs="0"/>
 *         &lt;element name="workGroup" type="{http://employee.ellipse.enterpriseservice.mincom.com}workGroup" minOccurs="0"/>
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
@XmlType(name = "EmployeeServiceRetrieveEmpForReqmtsReplyDTO", propOrder = {
    "activity",
    "activityDescription",
    "courseMatch",
    "courseMatchedWeightedPercentage",
    "courseNumberMatched",
    "courseTotalScore",
    "employee",
    "employeeFormattedName",
    "jobClassLevel",
    "jobClassLevelDesc",
    "matchedWeightedPercentage",
    "numberMatched",
    "payrollEmployeeInd",
    "persEmpStatusDesc",
    "personnelStatusDesc",
    "physicalLocation",
    "physicalLocationDesc",
    "position",
    "positionDesc",
    "positionMatch",
    "positionMatchedWeightedPercentage",
    "positionNumberMatched",
    "positionTotalScore",
    "resourceCompetency",
    "resourceCompetencyDesc",
    "resourceMatch",
    "resourceMatchedWeightedPercentage",
    "resourceNumberMatched",
    "resourceTotalScore",
    "totalScore",
    "workGroup",
    "workGroupDesc",
    "workLocation",
    "workLocationDesc"
})
public class EmployeeServiceRetrieveEmpForReqmtsReplyDTO
    extends AbstractReplyDTO
{

    protected String activity;
    protected String activityDescription;
    protected ArrayOfString courseMatch;
    protected BigDecimal courseMatchedWeightedPercentage;
    protected BigDecimal courseNumberMatched;
    protected BigDecimal courseTotalScore;
    protected String employee;
    protected String employeeFormattedName;
    protected String jobClassLevel;
    protected String jobClassLevelDesc;
    protected BigDecimal matchedWeightedPercentage;
    protected BigDecimal numberMatched;
    protected Boolean payrollEmployeeInd;
    protected String persEmpStatusDesc;
    protected String personnelStatusDesc;
    protected String physicalLocation;
    protected String physicalLocationDesc;
    protected String position;
    protected String positionDesc;
    protected ArrayOfString positionMatch;
    protected BigDecimal positionMatchedWeightedPercentage;
    protected BigDecimal positionNumberMatched;
    protected BigDecimal positionTotalScore;
    protected ArrayOfString resourceCompetency;
    protected ArrayOfString resourceCompetencyDesc;
    protected ArrayOfString resourceMatch;
    protected BigDecimal resourceMatchedWeightedPercentage;
    protected BigDecimal resourceNumberMatched;
    protected BigDecimal resourceTotalScore;
    protected BigDecimal totalScore;
    protected String workGroup;
    protected String workGroupDesc;
    protected String workLocation;
    protected String workLocationDesc;

    /**
     * Gets the value of the activity property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getActivity() {
        return activity;
    }

    /**
     * Sets the value of the activity property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setActivity(String value) {
        this.activity = value;
    }

    /**
     * Gets the value of the activityDescription property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getActivityDescription() {
        return activityDescription;
    }

    /**
     * Sets the value of the activityDescription property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setActivityDescription(String value) {
        this.activityDescription = value;
    }

    /**
     * Gets the value of the courseMatch property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfString }
     *     
     */
    public ArrayOfString getCourseMatch() {
        return courseMatch;
    }

    /**
     * Sets the value of the courseMatch property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfString }
     *     
     */
    public void setCourseMatch(ArrayOfString value) {
        this.courseMatch = value;
    }

    /**
     * Gets the value of the courseMatchedWeightedPercentage property.
     * 
     * @return
     *     possible object is
     *     {@link BigDecimal }
     *     
     */
    public BigDecimal getCourseMatchedWeightedPercentage() {
        return courseMatchedWeightedPercentage;
    }

    /**
     * Sets the value of the courseMatchedWeightedPercentage property.
     * 
     * @param value
     *     allowed object is
     *     {@link BigDecimal }
     *     
     */
    public void setCourseMatchedWeightedPercentage(BigDecimal value) {
        this.courseMatchedWeightedPercentage = value;
    }

    /**
     * Gets the value of the courseNumberMatched property.
     * 
     * @return
     *     possible object is
     *     {@link BigDecimal }
     *     
     */
    public BigDecimal getCourseNumberMatched() {
        return courseNumberMatched;
    }

    /**
     * Sets the value of the courseNumberMatched property.
     * 
     * @param value
     *     allowed object is
     *     {@link BigDecimal }
     *     
     */
    public void setCourseNumberMatched(BigDecimal value) {
        this.courseNumberMatched = value;
    }

    /**
     * Gets the value of the courseTotalScore property.
     * 
     * @return
     *     possible object is
     *     {@link BigDecimal }
     *     
     */
    public BigDecimal getCourseTotalScore() {
        return courseTotalScore;
    }

    /**
     * Sets the value of the courseTotalScore property.
     * 
     * @param value
     *     allowed object is
     *     {@link BigDecimal }
     *     
     */
    public void setCourseTotalScore(BigDecimal value) {
        this.courseTotalScore = value;
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
     * Gets the value of the jobClassLevel property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getJobClassLevel() {
        return jobClassLevel;
    }

    /**
     * Sets the value of the jobClassLevel property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setJobClassLevel(String value) {
        this.jobClassLevel = value;
    }

    /**
     * Gets the value of the jobClassLevelDesc property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getJobClassLevelDesc() {
        return jobClassLevelDesc;
    }

    /**
     * Sets the value of the jobClassLevelDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setJobClassLevelDesc(String value) {
        this.jobClassLevelDesc = value;
    }

    /**
     * Gets the value of the matchedWeightedPercentage property.
     * 
     * @return
     *     possible object is
     *     {@link BigDecimal }
     *     
     */
    public BigDecimal getMatchedWeightedPercentage() {
        return matchedWeightedPercentage;
    }

    /**
     * Sets the value of the matchedWeightedPercentage property.
     * 
     * @param value
     *     allowed object is
     *     {@link BigDecimal }
     *     
     */
    public void setMatchedWeightedPercentage(BigDecimal value) {
        this.matchedWeightedPercentage = value;
    }

    /**
     * Gets the value of the numberMatched property.
     * 
     * @return
     *     possible object is
     *     {@link BigDecimal }
     *     
     */
    public BigDecimal getNumberMatched() {
        return numberMatched;
    }

    /**
     * Sets the value of the numberMatched property.
     * 
     * @param value
     *     allowed object is
     *     {@link BigDecimal }
     *     
     */
    public void setNumberMatched(BigDecimal value) {
        this.numberMatched = value;
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
     * Gets the value of the positionMatch property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfString }
     *     
     */
    public ArrayOfString getPositionMatch() {
        return positionMatch;
    }

    /**
     * Sets the value of the positionMatch property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfString }
     *     
     */
    public void setPositionMatch(ArrayOfString value) {
        this.positionMatch = value;
    }

    /**
     * Gets the value of the positionMatchedWeightedPercentage property.
     * 
     * @return
     *     possible object is
     *     {@link BigDecimal }
     *     
     */
    public BigDecimal getPositionMatchedWeightedPercentage() {
        return positionMatchedWeightedPercentage;
    }

    /**
     * Sets the value of the positionMatchedWeightedPercentage property.
     * 
     * @param value
     *     allowed object is
     *     {@link BigDecimal }
     *     
     */
    public void setPositionMatchedWeightedPercentage(BigDecimal value) {
        this.positionMatchedWeightedPercentage = value;
    }

    /**
     * Gets the value of the positionNumberMatched property.
     * 
     * @return
     *     possible object is
     *     {@link BigDecimal }
     *     
     */
    public BigDecimal getPositionNumberMatched() {
        return positionNumberMatched;
    }

    /**
     * Sets the value of the positionNumberMatched property.
     * 
     * @param value
     *     allowed object is
     *     {@link BigDecimal }
     *     
     */
    public void setPositionNumberMatched(BigDecimal value) {
        this.positionNumberMatched = value;
    }

    /**
     * Gets the value of the positionTotalScore property.
     * 
     * @return
     *     possible object is
     *     {@link BigDecimal }
     *     
     */
    public BigDecimal getPositionTotalScore() {
        return positionTotalScore;
    }

    /**
     * Sets the value of the positionTotalScore property.
     * 
     * @param value
     *     allowed object is
     *     {@link BigDecimal }
     *     
     */
    public void setPositionTotalScore(BigDecimal value) {
        this.positionTotalScore = value;
    }

    /**
     * Gets the value of the resourceCompetency property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfString }
     *     
     */
    public ArrayOfString getResourceCompetency() {
        return resourceCompetency;
    }

    /**
     * Sets the value of the resourceCompetency property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfString }
     *     
     */
    public void setResourceCompetency(ArrayOfString value) {
        this.resourceCompetency = value;
    }

    /**
     * Gets the value of the resourceCompetencyDesc property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfString }
     *     
     */
    public ArrayOfString getResourceCompetencyDesc() {
        return resourceCompetencyDesc;
    }

    /**
     * Sets the value of the resourceCompetencyDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfString }
     *     
     */
    public void setResourceCompetencyDesc(ArrayOfString value) {
        this.resourceCompetencyDesc = value;
    }

    /**
     * Gets the value of the resourceMatch property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfString }
     *     
     */
    public ArrayOfString getResourceMatch() {
        return resourceMatch;
    }

    /**
     * Sets the value of the resourceMatch property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfString }
     *     
     */
    public void setResourceMatch(ArrayOfString value) {
        this.resourceMatch = value;
    }

    /**
     * Gets the value of the resourceMatchedWeightedPercentage property.
     * 
     * @return
     *     possible object is
     *     {@link BigDecimal }
     *     
     */
    public BigDecimal getResourceMatchedWeightedPercentage() {
        return resourceMatchedWeightedPercentage;
    }

    /**
     * Sets the value of the resourceMatchedWeightedPercentage property.
     * 
     * @param value
     *     allowed object is
     *     {@link BigDecimal }
     *     
     */
    public void setResourceMatchedWeightedPercentage(BigDecimal value) {
        this.resourceMatchedWeightedPercentage = value;
    }

    /**
     * Gets the value of the resourceNumberMatched property.
     * 
     * @return
     *     possible object is
     *     {@link BigDecimal }
     *     
     */
    public BigDecimal getResourceNumberMatched() {
        return resourceNumberMatched;
    }

    /**
     * Sets the value of the resourceNumberMatched property.
     * 
     * @param value
     *     allowed object is
     *     {@link BigDecimal }
     *     
     */
    public void setResourceNumberMatched(BigDecimal value) {
        this.resourceNumberMatched = value;
    }

    /**
     * Gets the value of the resourceTotalScore property.
     * 
     * @return
     *     possible object is
     *     {@link BigDecimal }
     *     
     */
    public BigDecimal getResourceTotalScore() {
        return resourceTotalScore;
    }

    /**
     * Sets the value of the resourceTotalScore property.
     * 
     * @param value
     *     allowed object is
     *     {@link BigDecimal }
     *     
     */
    public void setResourceTotalScore(BigDecimal value) {
        this.resourceTotalScore = value;
    }

    /**
     * Gets the value of the totalScore property.
     * 
     * @return
     *     possible object is
     *     {@link BigDecimal }
     *     
     */
    public BigDecimal getTotalScore() {
        return totalScore;
    }

    /**
     * Sets the value of the totalScore property.
     * 
     * @param value
     *     allowed object is
     *     {@link BigDecimal }
     *     
     */
    public void setTotalScore(BigDecimal value) {
        this.totalScore = value;
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
