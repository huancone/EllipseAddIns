
package com.mincom.enterpriseservice.ellipse.employee;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.AbstractDTO;


/**
 * <p>Java class for EmployeeServiceFetchLeaveHeaderRequestDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="EmployeeServiceFetchLeaveHeaderRequestDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractDTO">
 *       &lt;sequence>
 *         &lt;element name="employee" type="{http://employee.ellipse.enterpriseservice.mincom.com}employee" minOccurs="0"/>
 *         &lt;element name="lastUsedEmpInd" type="{http://employee.ellipse.enterpriseservice.mincom.com}lastUsedEmpInd" minOccurs="0"/>
 *         &lt;element name="requiredAttributes" type="{http://employee.ellipse.enterpriseservice.mincom.com}EmployeeServiceFetchLeaveHeaderRequiredAttributesDTO" minOccurs="0"/>
 *         &lt;element name="workGroupStartDate" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/extension>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "EmployeeServiceFetchLeaveHeaderRequestDTO", propOrder = {
    "employee",
    "lastUsedEmpInd",
    "requiredAttributes",
    "workGroupStartDate"
})
public class EmployeeServiceFetchLeaveHeaderRequestDTO
    extends AbstractDTO
{

    protected String employee;
    protected Boolean lastUsedEmpInd;
    protected EmployeeServiceFetchLeaveHeaderRequiredAttributesDTO requiredAttributes;
    protected String workGroupStartDate;

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
     * Gets the value of the lastUsedEmpInd property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isLastUsedEmpInd() {
        return lastUsedEmpInd;
    }

    /**
     * Sets the value of the lastUsedEmpInd property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setLastUsedEmpInd(Boolean value) {
        this.lastUsedEmpInd = value;
    }

    /**
     * Gets the value of the requiredAttributes property.
     * 
     * @return
     *     possible object is
     *     {@link EmployeeServiceFetchLeaveHeaderRequiredAttributesDTO }
     *     
     */
    public EmployeeServiceFetchLeaveHeaderRequiredAttributesDTO getRequiredAttributes() {
        return requiredAttributes;
    }

    /**
     * Sets the value of the requiredAttributes property.
     * 
     * @param value
     *     allowed object is
     *     {@link EmployeeServiceFetchLeaveHeaderRequiredAttributesDTO }
     *     
     */
    public void setRequiredAttributes(EmployeeServiceFetchLeaveHeaderRequiredAttributesDTO value) {
        this.requiredAttributes = value;
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

}
