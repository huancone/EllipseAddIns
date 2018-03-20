
package com.mincom.enterpriseservice.ellipse.employee;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.AbstractRequiredAttributesDTO;


/**
 * <p>Java class for EmployeeServiceTransferPositionRequiredAttributesDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="EmployeeServiceTransferPositionRequiredAttributesDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractRequiredAttributesDTO">
 *       &lt;sequence>
 *         &lt;element name="returnActualFTEPercent" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnAuthorityPercent" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnDataReferenceNo" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnDeathDate" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnDeathReason" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnDeathReasonDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnEmployee" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnEmployeeFormattedName" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnExitType" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnGlobalProfile" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelStatus" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelStatusDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPosition" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPositionDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPositionEndDate" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPositionReason" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPositionReasonDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPositionStartDate" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPrimRepCode" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPrimRepCodeDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnTransferStartDate" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnTransferType" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnTransferTypeDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/extension>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "EmployeeServiceTransferPositionRequiredAttributesDTO", propOrder = {
    "returnActualFTEPercent",
    "returnAuthorityPercent",
    "returnDataReferenceNo",
    "returnDeathDate",
    "returnDeathReason",
    "returnDeathReasonDesc",
    "returnEmployee",
    "returnEmployeeFormattedName",
    "returnExitType",
    "returnGlobalProfile",
    "returnPersonnelStatus",
    "returnPersonnelStatusDesc",
    "returnPosition",
    "returnPositionDesc",
    "returnPositionEndDate",
    "returnPositionReason",
    "returnPositionReasonDesc",
    "returnPositionStartDate",
    "returnPrimRepCode",
    "returnPrimRepCodeDesc",
    "returnTransferStartDate",
    "returnTransferType",
    "returnTransferTypeDesc"
})
public class EmployeeServiceTransferPositionRequiredAttributesDTO
    extends AbstractRequiredAttributesDTO
{

    protected Boolean returnActualFTEPercent;
    protected Boolean returnAuthorityPercent;
    protected Boolean returnDataReferenceNo;
    protected Boolean returnDeathDate;
    protected Boolean returnDeathReason;
    protected Boolean returnDeathReasonDesc;
    protected Boolean returnEmployee;
    protected Boolean returnEmployeeFormattedName;
    protected Boolean returnExitType;
    protected Boolean returnGlobalProfile;
    protected Boolean returnPersonnelStatus;
    protected Boolean returnPersonnelStatusDesc;
    protected Boolean returnPosition;
    protected Boolean returnPositionDesc;
    protected Boolean returnPositionEndDate;
    protected Boolean returnPositionReason;
    protected Boolean returnPositionReasonDesc;
    protected Boolean returnPositionStartDate;
    protected Boolean returnPrimRepCode;
    protected Boolean returnPrimRepCodeDesc;
    protected Boolean returnTransferStartDate;
    protected Boolean returnTransferType;
    protected Boolean returnTransferTypeDesc;

    /**
     * Gets the value of the returnActualFTEPercent property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnActualFTEPercent() {
        return returnActualFTEPercent;
    }

    /**
     * Sets the value of the returnActualFTEPercent property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnActualFTEPercent(Boolean value) {
        this.returnActualFTEPercent = value;
    }

    /**
     * Gets the value of the returnAuthorityPercent property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnAuthorityPercent() {
        return returnAuthorityPercent;
    }

    /**
     * Sets the value of the returnAuthorityPercent property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnAuthorityPercent(Boolean value) {
        this.returnAuthorityPercent = value;
    }

    /**
     * Gets the value of the returnDataReferenceNo property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnDataReferenceNo() {
        return returnDataReferenceNo;
    }

    /**
     * Sets the value of the returnDataReferenceNo property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnDataReferenceNo(Boolean value) {
        this.returnDataReferenceNo = value;
    }

    /**
     * Gets the value of the returnDeathDate property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnDeathDate() {
        return returnDeathDate;
    }

    /**
     * Sets the value of the returnDeathDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnDeathDate(Boolean value) {
        this.returnDeathDate = value;
    }

    /**
     * Gets the value of the returnDeathReason property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnDeathReason() {
        return returnDeathReason;
    }

    /**
     * Sets the value of the returnDeathReason property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnDeathReason(Boolean value) {
        this.returnDeathReason = value;
    }

    /**
     * Gets the value of the returnDeathReasonDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnDeathReasonDesc() {
        return returnDeathReasonDesc;
    }

    /**
     * Sets the value of the returnDeathReasonDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnDeathReasonDesc(Boolean value) {
        this.returnDeathReasonDesc = value;
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
     * Gets the value of the returnExitType property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnExitType() {
        return returnExitType;
    }

    /**
     * Sets the value of the returnExitType property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnExitType(Boolean value) {
        this.returnExitType = value;
    }

    /**
     * Gets the value of the returnGlobalProfile property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnGlobalProfile() {
        return returnGlobalProfile;
    }

    /**
     * Sets the value of the returnGlobalProfile property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnGlobalProfile(Boolean value) {
        this.returnGlobalProfile = value;
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
     * Gets the value of the returnPositionEndDate property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPositionEndDate() {
        return returnPositionEndDate;
    }

    /**
     * Sets the value of the returnPositionEndDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPositionEndDate(Boolean value) {
        this.returnPositionEndDate = value;
    }

    /**
     * Gets the value of the returnPositionReason property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPositionReason() {
        return returnPositionReason;
    }

    /**
     * Sets the value of the returnPositionReason property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPositionReason(Boolean value) {
        this.returnPositionReason = value;
    }

    /**
     * Gets the value of the returnPositionReasonDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPositionReasonDesc() {
        return returnPositionReasonDesc;
    }

    /**
     * Sets the value of the returnPositionReasonDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPositionReasonDesc(Boolean value) {
        this.returnPositionReasonDesc = value;
    }

    /**
     * Gets the value of the returnPositionStartDate property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPositionStartDate() {
        return returnPositionStartDate;
    }

    /**
     * Sets the value of the returnPositionStartDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPositionStartDate(Boolean value) {
        this.returnPositionStartDate = value;
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
     * Gets the value of the returnTransferStartDate property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnTransferStartDate() {
        return returnTransferStartDate;
    }

    /**
     * Sets the value of the returnTransferStartDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnTransferStartDate(Boolean value) {
        this.returnTransferStartDate = value;
    }

    /**
     * Gets the value of the returnTransferType property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnTransferType() {
        return returnTransferType;
    }

    /**
     * Sets the value of the returnTransferType property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnTransferType(Boolean value) {
        this.returnTransferType = value;
    }

    /**
     * Gets the value of the returnTransferTypeDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnTransferTypeDesc() {
        return returnTransferTypeDesc;
    }

    /**
     * Sets the value of the returnTransferTypeDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnTransferTypeDesc(Boolean value) {
        this.returnTransferTypeDesc = value;
    }

}
