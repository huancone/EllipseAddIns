
package com.mincom.enterpriseservice.ellipse.employee;

import java.util.ArrayList;
import java.util.List;
import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlType;


/**
 * <p>Java class for ArrayOfEmployeeServiceFetchLeaveHeaderRequestDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="ArrayOfEmployeeServiceFetchLeaveHeaderRequestDTO">
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *       &lt;sequence>
 *         &lt;element name="EmployeeServiceFetchLeaveHeaderRequestDTO" type="{http://employee.ellipse.enterpriseservice.mincom.com}EmployeeServiceFetchLeaveHeaderRequestDTO" maxOccurs="unbounded" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "ArrayOfEmployeeServiceFetchLeaveHeaderRequestDTO", propOrder = {
    "employeeServiceFetchLeaveHeaderRequestDTO"
})
public class ArrayOfEmployeeServiceFetchLeaveHeaderRequestDTO {

    @XmlElement(name = "EmployeeServiceFetchLeaveHeaderRequestDTO", nillable = true)
    protected List<EmployeeServiceFetchLeaveHeaderRequestDTO> employeeServiceFetchLeaveHeaderRequestDTO;

    /**
     * Gets the value of the employeeServiceFetchLeaveHeaderRequestDTO property.
     * 
     * <p>
     * This accessor method returns a reference to the live list,
     * not a snapshot. Therefore any modification you make to the
     * returned list will be present inside the JAXB object.
     * This is why there is not a <CODE>set</CODE> method for the employeeServiceFetchLeaveHeaderRequestDTO property.
     * 
     * <p>
     * For example, to add a new item, do as follows:
     * <pre>
     *    getEmployeeServiceFetchLeaveHeaderRequestDTO().add(newItem);
     * </pre>
     * 
     * 
     * <p>
     * Objects of the following type(s) are allowed in the list
     * {@link EmployeeServiceFetchLeaveHeaderRequestDTO }
     * 
     * 
     */
    public List<EmployeeServiceFetchLeaveHeaderRequestDTO> getEmployeeServiceFetchLeaveHeaderRequestDTO() {
        if (employeeServiceFetchLeaveHeaderRequestDTO == null) {
            employeeServiceFetchLeaveHeaderRequestDTO = new ArrayList<EmployeeServiceFetchLeaveHeaderRequestDTO>();
        }
        return this.employeeServiceFetchLeaveHeaderRequestDTO;
    }

}
