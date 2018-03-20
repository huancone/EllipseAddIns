
package com.mincom.enterpriseservice.ellipse.employee;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.AbstractDTO;


/**
 * <p>Java class for EmployeeServiceRetrieveEmpForReqmtsRequestDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="EmployeeServiceRetrieveEmpForReqmtsRequestDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractDTO">
 *       &lt;sequence>
 *         &lt;element name="activity" type="{http://employee.ellipse.enterpriseservice.mincom.com}activity" minOccurs="0"/>
 *         &lt;element name="employee" type="{http://employee.ellipse.enterpriseservice.mincom.com}employee" minOccurs="0"/>
 *         &lt;element name="physicalLocation" type="{http://employee.ellipse.enterpriseservice.mincom.com}physicalLocation" minOccurs="0"/>
 *         &lt;element name="position" type="{http://employee.ellipse.enterpriseservice.mincom.com}position" minOccurs="0"/>
 *         &lt;element name="reqmtCourseId" type="{http://employee.ellipse.enterpriseservice.mincom.com}ArrayOfString" minOccurs="0"/>
 *         &lt;element name="reqmtCourseMandatory" type="{http://employee.ellipse.enterpriseservice.mincom.com}ArrayOfString" minOccurs="0"/>
 *         &lt;element name="reqmtCourseWeight" type="{http://employee.ellipse.enterpriseservice.mincom.com}ArrayOfString" minOccurs="0"/>
 *         &lt;element name="reqmtPosition" type="{http://employee.ellipse.enterpriseservice.mincom.com}ArrayOfString" minOccurs="0"/>
 *         &lt;element name="reqmtPositionIndicator" type="{http://employee.ellipse.enterpriseservice.mincom.com}reqmtPositionIndicator" minOccurs="0"/>
 *         &lt;element name="reqmtPosnMandatory" type="{http://employee.ellipse.enterpriseservice.mincom.com}ArrayOfString" minOccurs="0"/>
 *         &lt;element name="reqmtPosnWeight" type="{http://employee.ellipse.enterpriseservice.mincom.com}ArrayOfString" minOccurs="0"/>
 *         &lt;element name="reqmtResourceClass" type="{http://employee.ellipse.enterpriseservice.mincom.com}ArrayOfString" minOccurs="0"/>
 *         &lt;element name="reqmtResourceCode" type="{http://employee.ellipse.enterpriseservice.mincom.com}ArrayOfString" minOccurs="0"/>
 *         &lt;element name="reqmtResourceCompetencyLevel" type="{http://employee.ellipse.enterpriseservice.mincom.com}ArrayOfString" minOccurs="0"/>
 *         &lt;element name="reqmtResourceIndicator" type="{http://employee.ellipse.enterpriseservice.mincom.com}reqmtResourceIndicator" minOccurs="0"/>
 *         &lt;element name="reqmtResourceMandatory" type="{http://employee.ellipse.enterpriseservice.mincom.com}ArrayOfString" minOccurs="0"/>
 *         &lt;element name="reqmtResourceWeight" type="{http://employee.ellipse.enterpriseservice.mincom.com}ArrayOfString" minOccurs="0"/>
 *         &lt;element name="searchStatus" type="{http://employee.ellipse.enterpriseservice.mincom.com}searchStatus" minOccurs="0"/>
 *         &lt;element name="workGroup" type="{http://employee.ellipse.enterpriseservice.mincom.com}workGroup" minOccurs="0"/>
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
@XmlType(name = "EmployeeServiceRetrieveEmpForReqmtsRequestDTO", propOrder = {
    "activity",
    "employee",
    "physicalLocation",
    "position",
    "reqmtCourseId",
    "reqmtCourseMandatory",
    "reqmtCourseWeight",
    "reqmtPosition",
    "reqmtPositionIndicator",
    "reqmtPosnMandatory",
    "reqmtPosnWeight",
    "reqmtResourceClass",
    "reqmtResourceCode",
    "reqmtResourceCompetencyLevel",
    "reqmtResourceIndicator",
    "reqmtResourceMandatory",
    "reqmtResourceWeight",
    "searchStatus",
    "workGroup",
    "workLocation"
})
public class EmployeeServiceRetrieveEmpForReqmtsRequestDTO
    extends AbstractDTO
{

    protected String activity;
    protected String employee;
    protected String physicalLocation;
    protected String position;
    protected ArrayOfString reqmtCourseId;
    protected ArrayOfString reqmtCourseMandatory;
    protected ArrayOfString reqmtCourseWeight;
    protected ArrayOfString reqmtPosition;
    protected String reqmtPositionIndicator;
    protected ArrayOfString reqmtPosnMandatory;
    protected ArrayOfString reqmtPosnWeight;
    protected ArrayOfString reqmtResourceClass;
    protected ArrayOfString reqmtResourceCode;
    protected ArrayOfString reqmtResourceCompetencyLevel;
    protected String reqmtResourceIndicator;
    protected ArrayOfString reqmtResourceMandatory;
    protected ArrayOfString reqmtResourceWeight;
    protected String searchStatus;
    protected String workGroup;
    protected String workLocation;

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
     * Gets the value of the reqmtCourseId property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfString }
     *     
     */
    public ArrayOfString getReqmtCourseId() {
        return reqmtCourseId;
    }

    /**
     * Sets the value of the reqmtCourseId property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfString }
     *     
     */
    public void setReqmtCourseId(ArrayOfString value) {
        this.reqmtCourseId = value;
    }

    /**
     * Gets the value of the reqmtCourseMandatory property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfString }
     *     
     */
    public ArrayOfString getReqmtCourseMandatory() {
        return reqmtCourseMandatory;
    }

    /**
     * Sets the value of the reqmtCourseMandatory property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfString }
     *     
     */
    public void setReqmtCourseMandatory(ArrayOfString value) {
        this.reqmtCourseMandatory = value;
    }

    /**
     * Gets the value of the reqmtCourseWeight property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfString }
     *     
     */
    public ArrayOfString getReqmtCourseWeight() {
        return reqmtCourseWeight;
    }

    /**
     * Sets the value of the reqmtCourseWeight property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfString }
     *     
     */
    public void setReqmtCourseWeight(ArrayOfString value) {
        this.reqmtCourseWeight = value;
    }

    /**
     * Gets the value of the reqmtPosition property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfString }
     *     
     */
    public ArrayOfString getReqmtPosition() {
        return reqmtPosition;
    }

    /**
     * Sets the value of the reqmtPosition property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfString }
     *     
     */
    public void setReqmtPosition(ArrayOfString value) {
        this.reqmtPosition = value;
    }

    /**
     * Gets the value of the reqmtPositionIndicator property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getReqmtPositionIndicator() {
        return reqmtPositionIndicator;
    }

    /**
     * Sets the value of the reqmtPositionIndicator property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setReqmtPositionIndicator(String value) {
        this.reqmtPositionIndicator = value;
    }

    /**
     * Gets the value of the reqmtPosnMandatory property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfString }
     *     
     */
    public ArrayOfString getReqmtPosnMandatory() {
        return reqmtPosnMandatory;
    }

    /**
     * Sets the value of the reqmtPosnMandatory property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfString }
     *     
     */
    public void setReqmtPosnMandatory(ArrayOfString value) {
        this.reqmtPosnMandatory = value;
    }

    /**
     * Gets the value of the reqmtPosnWeight property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfString }
     *     
     */
    public ArrayOfString getReqmtPosnWeight() {
        return reqmtPosnWeight;
    }

    /**
     * Sets the value of the reqmtPosnWeight property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfString }
     *     
     */
    public void setReqmtPosnWeight(ArrayOfString value) {
        this.reqmtPosnWeight = value;
    }

    /**
     * Gets the value of the reqmtResourceClass property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfString }
     *     
     */
    public ArrayOfString getReqmtResourceClass() {
        return reqmtResourceClass;
    }

    /**
     * Sets the value of the reqmtResourceClass property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfString }
     *     
     */
    public void setReqmtResourceClass(ArrayOfString value) {
        this.reqmtResourceClass = value;
    }

    /**
     * Gets the value of the reqmtResourceCode property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfString }
     *     
     */
    public ArrayOfString getReqmtResourceCode() {
        return reqmtResourceCode;
    }

    /**
     * Sets the value of the reqmtResourceCode property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfString }
     *     
     */
    public void setReqmtResourceCode(ArrayOfString value) {
        this.reqmtResourceCode = value;
    }

    /**
     * Gets the value of the reqmtResourceCompetencyLevel property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfString }
     *     
     */
    public ArrayOfString getReqmtResourceCompetencyLevel() {
        return reqmtResourceCompetencyLevel;
    }

    /**
     * Sets the value of the reqmtResourceCompetencyLevel property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfString }
     *     
     */
    public void setReqmtResourceCompetencyLevel(ArrayOfString value) {
        this.reqmtResourceCompetencyLevel = value;
    }

    /**
     * Gets the value of the reqmtResourceIndicator property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getReqmtResourceIndicator() {
        return reqmtResourceIndicator;
    }

    /**
     * Sets the value of the reqmtResourceIndicator property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setReqmtResourceIndicator(String value) {
        this.reqmtResourceIndicator = value;
    }

    /**
     * Gets the value of the reqmtResourceMandatory property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfString }
     *     
     */
    public ArrayOfString getReqmtResourceMandatory() {
        return reqmtResourceMandatory;
    }

    /**
     * Sets the value of the reqmtResourceMandatory property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfString }
     *     
     */
    public void setReqmtResourceMandatory(ArrayOfString value) {
        this.reqmtResourceMandatory = value;
    }

    /**
     * Gets the value of the reqmtResourceWeight property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfString }
     *     
     */
    public ArrayOfString getReqmtResourceWeight() {
        return reqmtResourceWeight;
    }

    /**
     * Sets the value of the reqmtResourceWeight property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfString }
     *     
     */
    public void setReqmtResourceWeight(ArrayOfString value) {
        this.reqmtResourceWeight = value;
    }

    /**
     * Gets the value of the searchStatus property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getSearchStatus() {
        return searchStatus;
    }

    /**
     * Sets the value of the searchStatus property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setSearchStatus(String value) {
        this.searchStatus = value;
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
