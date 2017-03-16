
package com.mincom.enterpriseservice.ellipse.employee;

import java.math.BigDecimal;
import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.AbstractDTO;


/**
 * <p>Java class for EmployeeServiceTransferPositionRequestDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="EmployeeServiceTransferPositionRequestDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractDTO">
 *       &lt;sequence>
 *         &lt;element name="bonaFideTermination" type="{http://employee.ellipse.enterpriseservice.mincom.com}bonaFideTermination" minOccurs="0"/>
 *         &lt;element name="copySalaryPackage" type="{http://employee.ellipse.enterpriseservice.mincom.com}copySalaryPackage" minOccurs="0"/>
 *         &lt;element name="dataReferenceNo" type="{http://employee.ellipse.enterpriseservice.mincom.com}dataReferenceNo" minOccurs="0"/>
 *         &lt;element name="deathDate" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *         &lt;element name="deathReason" type="{http://employee.ellipse.enterpriseservice.mincom.com}deathReason" minOccurs="0"/>
 *         &lt;element name="employee" type="{http://employee.ellipse.enterpriseservice.mincom.com}employee" minOccurs="0"/>
 *         &lt;element name="exitType" type="{http://employee.ellipse.enterpriseservice.mincom.com}exitType" minOccurs="0"/>
 *         &lt;element name="newActualFTEPercent" type="{http://employee.ellipse.enterpriseservice.mincom.com}newActualFTEPercent" minOccurs="0"/>
 *         &lt;element name="newAuthorityPercent" type="{http://employee.ellipse.enterpriseservice.mincom.com}newAuthorityPercent" minOccurs="0"/>
 *         &lt;element name="newGlobalProfile" type="{http://employee.ellipse.enterpriseservice.mincom.com}newGlobalProfile" minOccurs="0"/>
 *         &lt;element name="newPersonnelStatus" type="{http://employee.ellipse.enterpriseservice.mincom.com}newPersonnelStatus" minOccurs="0"/>
 *         &lt;element name="newPositionId" type="{http://employee.ellipse.enterpriseservice.mincom.com}newPositionId" minOccurs="0"/>
 *         &lt;element name="newPositionReason" type="{http://employee.ellipse.enterpriseservice.mincom.com}newPositionReason" minOccurs="0"/>
 *         &lt;element name="personnelStatus" type="{http://employee.ellipse.enterpriseservice.mincom.com}personnelStatus" minOccurs="0"/>
 *         &lt;element name="position" type="{http://employee.ellipse.enterpriseservice.mincom.com}position" minOccurs="0"/>
 *         &lt;element name="positionReason" type="{http://employee.ellipse.enterpriseservice.mincom.com}positionReason" minOccurs="0"/>
 *         &lt;element name="requiredAttributes" type="{http://employee.ellipse.enterpriseservice.mincom.com}EmployeeServiceTransferPositionRequiredAttributesDTO" minOccurs="0"/>
 *         &lt;element name="transferStartDate" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *         &lt;element name="transferType" type="{http://employee.ellipse.enterpriseservice.mincom.com}transferType" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/extension>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "EmployeeServiceTransferPositionRequestDTO", propOrder = {
    "bonaFideTermination",
    "copySalaryPackage",
    "dataReferenceNo",
    "deathDate",
    "deathReason",
    "employee",
    "exitType",
    "newActualFTEPercent",
    "newAuthorityPercent",
    "newGlobalProfile",
    "newPersonnelStatus",
    "newPositionId",
    "newPositionReason",
    "personnelStatus",
    "position",
    "positionReason",
    "requiredAttributes",
    "transferStartDate",
    "transferType"
})
public class EmployeeServiceTransferPositionRequestDTO
    extends AbstractDTO
{

    protected Boolean bonaFideTermination;
    protected Boolean copySalaryPackage;
    protected String dataReferenceNo;
    protected String deathDate;
    protected String deathReason;
    protected String employee;
    protected String exitType;
    protected BigDecimal newActualFTEPercent;
    protected BigDecimal newAuthorityPercent;
    protected String newGlobalProfile;
    protected String newPersonnelStatus;
    protected String newPositionId;
    protected String newPositionReason;
    protected String personnelStatus;
    protected String position;
    protected String positionReason;
    protected EmployeeServiceTransferPositionRequiredAttributesDTO requiredAttributes;
    protected String transferStartDate;
    protected String transferType;

    /**
     * Gets the value of the bonaFideTermination property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isBonaFideTermination() {
        return bonaFideTermination;
    }

    /**
     * Sets the value of the bonaFideTermination property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setBonaFideTermination(Boolean value) {
        this.bonaFideTermination = value;
    }

    /**
     * Gets the value of the copySalaryPackage property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isCopySalaryPackage() {
        return copySalaryPackage;
    }

    /**
     * Sets the value of the copySalaryPackage property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setCopySalaryPackage(Boolean value) {
        this.copySalaryPackage = value;
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
     * Gets the value of the newActualFTEPercent property.
     * 
     * @return
     *     possible object is
     *     {@link BigDecimal }
     *     
     */
    public BigDecimal getNewActualFTEPercent() {
        return newActualFTEPercent;
    }

    /**
     * Sets the value of the newActualFTEPercent property.
     * 
     * @param value
     *     allowed object is
     *     {@link BigDecimal }
     *     
     */
    public void setNewActualFTEPercent(BigDecimal value) {
        this.newActualFTEPercent = value;
    }

    /**
     * Gets the value of the newAuthorityPercent property.
     * 
     * @return
     *     possible object is
     *     {@link BigDecimal }
     *     
     */
    public BigDecimal getNewAuthorityPercent() {
        return newAuthorityPercent;
    }

    /**
     * Sets the value of the newAuthorityPercent property.
     * 
     * @param value
     *     allowed object is
     *     {@link BigDecimal }
     *     
     */
    public void setNewAuthorityPercent(BigDecimal value) {
        this.newAuthorityPercent = value;
    }

    /**
     * Gets the value of the newGlobalProfile property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getNewGlobalProfile() {
        return newGlobalProfile;
    }

    /**
     * Sets the value of the newGlobalProfile property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setNewGlobalProfile(String value) {
        this.newGlobalProfile = value;
    }

    /**
     * Gets the value of the newPersonnelStatus property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getNewPersonnelStatus() {
        return newPersonnelStatus;
    }

    /**
     * Sets the value of the newPersonnelStatus property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setNewPersonnelStatus(String value) {
        this.newPersonnelStatus = value;
    }

    /**
     * Gets the value of the newPositionId property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getNewPositionId() {
        return newPositionId;
    }

    /**
     * Sets the value of the newPositionId property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setNewPositionId(String value) {
        this.newPositionId = value;
    }

    /**
     * Gets the value of the newPositionReason property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getNewPositionReason() {
        return newPositionReason;
    }

    /**
     * Sets the value of the newPositionReason property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setNewPositionReason(String value) {
        this.newPositionReason = value;
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
     * Gets the value of the requiredAttributes property.
     * 
     * @return
     *     possible object is
     *     {@link EmployeeServiceTransferPositionRequiredAttributesDTO }
     *     
     */
    public EmployeeServiceTransferPositionRequiredAttributesDTO getRequiredAttributes() {
        return requiredAttributes;
    }

    /**
     * Sets the value of the requiredAttributes property.
     * 
     * @param value
     *     allowed object is
     *     {@link EmployeeServiceTransferPositionRequiredAttributesDTO }
     *     
     */
    public void setRequiredAttributes(EmployeeServiceTransferPositionRequiredAttributesDTO value) {
        this.requiredAttributes = value;
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

}
