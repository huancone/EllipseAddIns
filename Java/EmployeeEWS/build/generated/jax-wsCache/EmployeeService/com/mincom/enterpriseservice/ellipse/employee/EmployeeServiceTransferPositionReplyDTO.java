
package com.mincom.enterpriseservice.ellipse.employee;

import java.math.BigDecimal;
import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.AbstractReplyDTO;


/**
 * <p>Java class for EmployeeServiceTransferPositionReplyDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="EmployeeServiceTransferPositionReplyDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractReplyDTO">
 *       &lt;sequence>
 *         &lt;element name="actualFTEPercent" type="{http://employee.ellipse.enterpriseservice.mincom.com}actualFTEPercent" minOccurs="0"/>
 *         &lt;element name="authorityPercent" type="{http://employee.ellipse.enterpriseservice.mincom.com}authorityPercent" minOccurs="0"/>
 *         &lt;element name="dataReferenceNo" type="{http://employee.ellipse.enterpriseservice.mincom.com}dataReferenceNo" minOccurs="0"/>
 *         &lt;element name="deathDate" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *         &lt;element name="deathReason" type="{http://employee.ellipse.enterpriseservice.mincom.com}deathReason" minOccurs="0"/>
 *         &lt;element name="deathReasonDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}deathReasonDesc" minOccurs="0"/>
 *         &lt;element name="employee" type="{http://employee.ellipse.enterpriseservice.mincom.com}employee" minOccurs="0"/>
 *         &lt;element name="employeeFormattedName" type="{http://employee.ellipse.enterpriseservice.mincom.com}employeeFormattedName" minOccurs="0"/>
 *         &lt;element name="exitType" type="{http://employee.ellipse.enterpriseservice.mincom.com}exitType" minOccurs="0"/>
 *         &lt;element name="globalProfile" type="{http://employee.ellipse.enterpriseservice.mincom.com}globalProfile" minOccurs="0"/>
 *         &lt;element name="personnelStatus" type="{http://employee.ellipse.enterpriseservice.mincom.com}personnelStatus" minOccurs="0"/>
 *         &lt;element name="personnelStatusDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}personnelStatusDesc" minOccurs="0"/>
 *         &lt;element name="position" type="{http://employee.ellipse.enterpriseservice.mincom.com}position" minOccurs="0"/>
 *         &lt;element name="positionDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}positionDesc" minOccurs="0"/>
 *         &lt;element name="positionEndDate" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *         &lt;element name="positionReason" type="{http://employee.ellipse.enterpriseservice.mincom.com}positionReason" minOccurs="0"/>
 *         &lt;element name="positionReasonDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}positionReasonDesc" minOccurs="0"/>
 *         &lt;element name="positionStartDate" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *         &lt;element name="primRepCode" type="{http://employee.ellipse.enterpriseservice.mincom.com}primRepCode" minOccurs="0"/>
 *         &lt;element name="primRepCodeDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}primRepCodeDesc" minOccurs="0"/>
 *         &lt;element name="transferStartDate" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *         &lt;element name="transferType" type="{http://employee.ellipse.enterpriseservice.mincom.com}transferType" minOccurs="0"/>
 *         &lt;element name="transferTypeDesc" type="{http://employee.ellipse.enterpriseservice.mincom.com}transferTypeDesc" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/extension>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "EmployeeServiceTransferPositionReplyDTO", propOrder = {
    "actualFTEPercent",
    "authorityPercent",
    "dataReferenceNo",
    "deathDate",
    "deathReason",
    "deathReasonDesc",
    "employee",
    "employeeFormattedName",
    "exitType",
    "globalProfile",
    "personnelStatus",
    "personnelStatusDesc",
    "position",
    "positionDesc",
    "positionEndDate",
    "positionReason",
    "positionReasonDesc",
    "positionStartDate",
    "primRepCode",
    "primRepCodeDesc",
    "transferStartDate",
    "transferType",
    "transferTypeDesc"
})
public class EmployeeServiceTransferPositionReplyDTO
    extends AbstractReplyDTO
{

    protected BigDecimal actualFTEPercent;
    protected BigDecimal authorityPercent;
    protected String dataReferenceNo;
    protected String deathDate;
    protected String deathReason;
    protected String deathReasonDesc;
    protected String employee;
    protected String employeeFormattedName;
    protected String exitType;
    protected String globalProfile;
    protected String personnelStatus;
    protected String personnelStatusDesc;
    protected String position;
    protected String positionDesc;
    protected String positionEndDate;
    protected String positionReason;
    protected String positionReasonDesc;
    protected String positionStartDate;
    protected String primRepCode;
    protected String primRepCodeDesc;
    protected String transferStartDate;
    protected String transferType;
    protected String transferTypeDesc;

    /**
     * Gets the value of the actualFTEPercent property.
     * 
     * @return
     *     possible object is
     *     {@link BigDecimal }
     *     
     */
    public BigDecimal getActualFTEPercent() {
        return actualFTEPercent;
    }

    /**
     * Sets the value of the actualFTEPercent property.
     * 
     * @param value
     *     allowed object is
     *     {@link BigDecimal }
     *     
     */
    public void setActualFTEPercent(BigDecimal value) {
        this.actualFTEPercent = value;
    }

    /**
     * Gets the value of the authorityPercent property.
     * 
     * @return
     *     possible object is
     *     {@link BigDecimal }
     *     
     */
    public BigDecimal getAuthorityPercent() {
        return authorityPercent;
    }

    /**
     * Sets the value of the authorityPercent property.
     * 
     * @param value
     *     allowed object is
     *     {@link BigDecimal }
     *     
     */
    public void setAuthorityPercent(BigDecimal value) {
        this.authorityPercent = value;
    }

    /**
     * Gets the value of the dataReferenceNo property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getDataReferenceNo() {
        return dataReferenceNo;
    }

    /**
     * Sets the value of the dataReferenceNo property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setDataReferenceNo(String value) {
        this.dataReferenceNo = value;
    }

    /**
     * Gets the value of the deathDate property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getDeathDate() {
        return deathDate;
    }

    /**
     * Sets the value of the deathDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setDeathDate(String value) {
        this.deathDate = value;
    }

    /**
     * Gets the value of the deathReason property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getDeathReason() {
        return deathReason;
    }

    /**
     * Sets the value of the deathReason property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setDeathReason(String value) {
        this.deathReason = value;
    }

    /**
     * Gets the value of the deathReasonDesc property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getDeathReasonDesc() {
        return deathReasonDesc;
    }

    /**
     * Sets the value of the deathReasonDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setDeathReasonDesc(String value) {
        this.deathReasonDesc = value;
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
     * Gets the value of the exitType property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getExitType() {
        return exitType;
    }

    /**
     * Sets the value of the exitType property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setExitType(String value) {
        this.exitType = value;
    }

    /**
     * Gets the value of the globalProfile property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getGlobalProfile() {
        return globalProfile;
    }

    /**
     * Sets the value of the globalProfile property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setGlobalProfile(String value) {
        this.globalProfile = value;
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
     * Gets the value of the positionEndDate property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPositionEndDate() {
        return positionEndDate;
    }

    /**
     * Sets the value of the positionEndDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPositionEndDate(String value) {
        this.positionEndDate = value;
    }

    /**
     * Gets the value of the positionReason property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPositionReason() {
        return positionReason;
    }

    /**
     * Sets the value of the positionReason property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPositionReason(String value) {
        this.positionReason = value;
    }

    /**
     * Gets the value of the positionReasonDesc property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPositionReasonDesc() {
        return positionReasonDesc;
    }

    /**
     * Sets the value of the positionReasonDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPositionReasonDesc(String value) {
        this.positionReasonDesc = value;
    }

    /**
     * Gets the value of the positionStartDate property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPositionStartDate() {
        return positionStartDate;
    }

    /**
     * Sets the value of the positionStartDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPositionStartDate(String value) {
        this.positionStartDate = value;
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
     * Gets the value of the transferStartDate property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getTransferStartDate() {
        return transferStartDate;
    }

    /**
     * Sets the value of the transferStartDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setTransferStartDate(String value) {
        this.transferStartDate = value;
    }

    /**
     * Gets the value of the transferType property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getTransferType() {
        return transferType;
    }

    /**
     * Sets the value of the transferType property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setTransferType(String value) {
        this.transferType = value;
    }

    /**
     * Gets the value of the transferTypeDesc property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getTransferTypeDesc() {
        return transferTypeDesc;
    }

    /**
     * Sets the value of the transferTypeDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setTransferTypeDesc(String value) {
        this.transferTypeDesc = value;
    }

}
