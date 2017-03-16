
package com.mincom.enterpriseservice.ellipse.employee;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.AbstractRequiredAttributesDTO;


/**
 * <p>Java class for EmployeeServiceRetrieveEmpForReqmtsRequiredAttributesDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="EmployeeServiceRetrieveEmpForReqmtsRequiredAttributesDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractRequiredAttributesDTO">
 *       &lt;sequence>
 *         &lt;element name="returnActivity" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnActivityDescription" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnCourseMatch" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnCourseMatchedWeightedPercentage" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnCourseNumberMatched" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnCourseTotalScore" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnEmployee" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnEmployeeFormattedName" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnJobClassLevel" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnJobClassLevelDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnMatchedWeightedPercentage" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnNumberMatched" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPayrollEmployeeInd" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersEmpStatusDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelStatusDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPhysicalLocation" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPhysicalLocationDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPosition" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPositionDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPositionMatch" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPositionMatchedWeightedPercentage" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPositionNumberMatched" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPositionTotalScore" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnResourceCompetency" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnResourceCompetencyDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnResourceMatch" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnResourceMatchedWeightedPercentage" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnResourceNumberMatched" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnResourceTotalScore" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnTotalScore" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnWorkGroup" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
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
@XmlType(name = "EmployeeServiceRetrieveEmpForReqmtsRequiredAttributesDTO", propOrder = {
    "returnActivity",
    "returnActivityDescription",
    "returnCourseMatch",
    "returnCourseMatchedWeightedPercentage",
    "returnCourseNumberMatched",
    "returnCourseTotalScore",
    "returnEmployee",
    "returnEmployeeFormattedName",
    "returnJobClassLevel",
    "returnJobClassLevelDesc",
    "returnMatchedWeightedPercentage",
    "returnNumberMatched",
    "returnPayrollEmployeeInd",
    "returnPersEmpStatusDesc",
    "returnPersonnelStatusDesc",
    "returnPhysicalLocation",
    "returnPhysicalLocationDesc",
    "returnPosition",
    "returnPositionDesc",
    "returnPositionMatch",
    "returnPositionMatchedWeightedPercentage",
    "returnPositionNumberMatched",
    "returnPositionTotalScore",
    "returnResourceCompetency",
    "returnResourceCompetencyDesc",
    "returnResourceMatch",
    "returnResourceMatchedWeightedPercentage",
    "returnResourceNumberMatched",
    "returnResourceTotalScore",
    "returnTotalScore",
    "returnWorkGroup",
    "returnWorkGroupDesc",
    "returnWorkLocation",
    "returnWorkLocationDesc"
})
public class EmployeeServiceRetrieveEmpForReqmtsRequiredAttributesDTO
    extends AbstractRequiredAttributesDTO
{

    protected Boolean returnActivity;
    protected Boolean returnActivityDescription;
    protected Boolean returnCourseMatch;
    protected Boolean returnCourseMatchedWeightedPercentage;
    protected Boolean returnCourseNumberMatched;
    protected Boolean returnCourseTotalScore;
    protected Boolean returnEmployee;
    protected Boolean returnEmployeeFormattedName;
    protected Boolean returnJobClassLevel;
    protected Boolean returnJobClassLevelDesc;
    protected Boolean returnMatchedWeightedPercentage;
    protected Boolean returnNumberMatched;
    protected Boolean returnPayrollEmployeeInd;
    protected Boolean returnPersEmpStatusDesc;
    protected Boolean returnPersonnelStatusDesc;
    protected Boolean returnPhysicalLocation;
    protected Boolean returnPhysicalLocationDesc;
    protected Boolean returnPosition;
    protected Boolean returnPositionDesc;
    protected Boolean returnPositionMatch;
    protected Boolean returnPositionMatchedWeightedPercentage;
    protected Boolean returnPositionNumberMatched;
    protected Boolean returnPositionTotalScore;
    protected Boolean returnResourceCompetency;
    protected Boolean returnResourceCompetencyDesc;
    protected Boolean returnResourceMatch;
    protected Boolean returnResourceMatchedWeightedPercentage;
    protected Boolean returnResourceNumberMatched;
    protected Boolean returnResourceTotalScore;
    protected Boolean returnTotalScore;
    protected Boolean returnWorkGroup;
    protected Boolean returnWorkGroupDesc;
    protected Boolean returnWorkLocation;
    protected Boolean returnWorkLocationDesc;

    /**
     * Gets the value of the returnActivity property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnActivity() {
        return returnActivity;
    }

    /**
     * Sets the value of the returnActivity property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnActivity(Boolean value) {
        this.returnActivity = value;
    }

    /**
     * Gets the value of the returnActivityDescription property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnActivityDescription() {
        return returnActivityDescription;
    }

    /**
     * Sets the value of the returnActivityDescription property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnActivityDescription(Boolean value) {
        this.returnActivityDescription = value;
    }

    /**
     * Gets the value of the returnCourseMatch property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnCourseMatch() {
        return returnCourseMatch;
    }

    /**
     * Sets the value of the returnCourseMatch property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnCourseMatch(Boolean value) {
        this.returnCourseMatch = value;
    }

    /**
     * Gets the value of the returnCourseMatchedWeightedPercentage property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnCourseMatchedWeightedPercentage() {
        return returnCourseMatchedWeightedPercentage;
    }

    /**
     * Sets the value of the returnCourseMatchedWeightedPercentage property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnCourseMatchedWeightedPercentage(Boolean value) {
        this.returnCourseMatchedWeightedPercentage = value;
    }

    /**
     * Gets the value of the returnCourseNumberMatched property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnCourseNumberMatched() {
        return returnCourseNumberMatched;
    }

    /**
     * Sets the value of the returnCourseNumberMatched property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnCourseNumberMatched(Boolean value) {
        this.returnCourseNumberMatched = value;
    }

    /**
     * Gets the value of the returnCourseTotalScore property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnCourseTotalScore() {
        return returnCourseTotalScore;
    }

    /**
     * Sets the value of the returnCourseTotalScore property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnCourseTotalScore(Boolean value) {
        this.returnCourseTotalScore = value;
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
     * Gets the value of the returnJobClassLevel property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnJobClassLevel() {
        return returnJobClassLevel;
    }

    /**
     * Sets the value of the returnJobClassLevel property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnJobClassLevel(Boolean value) {
        this.returnJobClassLevel = value;
    }

    /**
     * Gets the value of the returnJobClassLevelDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnJobClassLevelDesc() {
        return returnJobClassLevelDesc;
    }

    /**
     * Sets the value of the returnJobClassLevelDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnJobClassLevelDesc(Boolean value) {
        this.returnJobClassLevelDesc = value;
    }

    /**
     * Gets the value of the returnMatchedWeightedPercentage property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnMatchedWeightedPercentage() {
        return returnMatchedWeightedPercentage;
    }

    /**
     * Sets the value of the returnMatchedWeightedPercentage property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnMatchedWeightedPercentage(Boolean value) {
        this.returnMatchedWeightedPercentage = value;
    }

    /**
     * Gets the value of the returnNumberMatched property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnNumberMatched() {
        return returnNumberMatched;
    }

    /**
     * Sets the value of the returnNumberMatched property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnNumberMatched(Boolean value) {
        this.returnNumberMatched = value;
    }

    /**
     * Gets the value of the returnPayrollEmployeeInd property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPayrollEmployeeInd() {
        return returnPayrollEmployeeInd;
    }

    /**
     * Sets the value of the returnPayrollEmployeeInd property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPayrollEmployeeInd(Boolean value) {
        this.returnPayrollEmployeeInd = value;
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
     * Gets the value of the returnPositionMatch property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPositionMatch() {
        return returnPositionMatch;
    }

    /**
     * Sets the value of the returnPositionMatch property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPositionMatch(Boolean value) {
        this.returnPositionMatch = value;
    }

    /**
     * Gets the value of the returnPositionMatchedWeightedPercentage property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPositionMatchedWeightedPercentage() {
        return returnPositionMatchedWeightedPercentage;
    }

    /**
     * Sets the value of the returnPositionMatchedWeightedPercentage property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPositionMatchedWeightedPercentage(Boolean value) {
        this.returnPositionMatchedWeightedPercentage = value;
    }

    /**
     * Gets the value of the returnPositionNumberMatched property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPositionNumberMatched() {
        return returnPositionNumberMatched;
    }

    /**
     * Sets the value of the returnPositionNumberMatched property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPositionNumberMatched(Boolean value) {
        this.returnPositionNumberMatched = value;
    }

    /**
     * Gets the value of the returnPositionTotalScore property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPositionTotalScore() {
        return returnPositionTotalScore;
    }

    /**
     * Sets the value of the returnPositionTotalScore property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPositionTotalScore(Boolean value) {
        this.returnPositionTotalScore = value;
    }

    /**
     * Gets the value of the returnResourceCompetency property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnResourceCompetency() {
        return returnResourceCompetency;
    }

    /**
     * Sets the value of the returnResourceCompetency property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnResourceCompetency(Boolean value) {
        this.returnResourceCompetency = value;
    }

    /**
     * Gets the value of the returnResourceCompetencyDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnResourceCompetencyDesc() {
        return returnResourceCompetencyDesc;
    }

    /**
     * Sets the value of the returnResourceCompetencyDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnResourceCompetencyDesc(Boolean value) {
        this.returnResourceCompetencyDesc = value;
    }

    /**
     * Gets the value of the returnResourceMatch property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnResourceMatch() {
        return returnResourceMatch;
    }

    /**
     * Sets the value of the returnResourceMatch property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnResourceMatch(Boolean value) {
        this.returnResourceMatch = value;
    }

    /**
     * Gets the value of the returnResourceMatchedWeightedPercentage property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnResourceMatchedWeightedPercentage() {
        return returnResourceMatchedWeightedPercentage;
    }

    /**
     * Sets the value of the returnResourceMatchedWeightedPercentage property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnResourceMatchedWeightedPercentage(Boolean value) {
        this.returnResourceMatchedWeightedPercentage = value;
    }

    /**
     * Gets the value of the returnResourceNumberMatched property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnResourceNumberMatched() {
        return returnResourceNumberMatched;
    }

    /**
     * Sets the value of the returnResourceNumberMatched property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnResourceNumberMatched(Boolean value) {
        this.returnResourceNumberMatched = value;
    }

    /**
     * Gets the value of the returnResourceTotalScore property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnResourceTotalScore() {
        return returnResourceTotalScore;
    }

    /**
     * Sets the value of the returnResourceTotalScore property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnResourceTotalScore(Boolean value) {
        this.returnResourceTotalScore = value;
    }

    /**
     * Gets the value of the returnTotalScore property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnTotalScore() {
        return returnTotalScore;
    }

    /**
     * Sets the value of the returnTotalScore property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnTotalScore(Boolean value) {
        this.returnTotalScore = value;
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
