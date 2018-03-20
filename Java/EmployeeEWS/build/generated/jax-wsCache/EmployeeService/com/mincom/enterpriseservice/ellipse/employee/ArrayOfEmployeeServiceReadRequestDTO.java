
package com.mincom.enterpriseservice.ellipse.employee;

import java.util.ArrayList;
import java.util.List;
import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlType;


/**
 * <p>Java class for ArrayOfEmployeeServiceReadRequestDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="ArrayOfEmployeeServiceReadRequestDTO">
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *       &lt;sequence>
 *         &lt;element name="EmployeeServiceReadRequestDTO" type="{http://employee.ellipse.enterpriseservice.mincom.com}EmployeeServiceReadRequestDTO" maxOccurs="unbounded" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "ArrayOfEmployeeServiceReadRequestDTO", propOrder = {
    "employeeServiceReadRequestDTO"
})
public class ArrayOfEmployeeServiceReadRequestDTO {

    @XmlElement(name = "EmployeeServiceReadRequestDTO", nillable = true)
    protected List<EmployeeServiceReadRequestDTO> employeeServiceReadRequestDTO;

    /**
     * Gets the value of the employeeServiceReadRequestDTO property.
     * 
     * <p>
     * This accessor method returns a reference to the live list,
     * not a snapshot. Therefore any modification you make to the
     * returned list will be present inside the JAXB object.
     * This is why there is not a <CODE>set</CODE> method for the employeeServiceReadRequestDTO property.
     * 
     * <p>
     * For example, to add a new item, do as follows:
     * <pre>
     *    getEmployeeServiceReadRequestDTO().add(newItem);
     * </pre>
     * 
     * 
     * <p>
     * Objects of the following type(s) are allowed in the list
     * {@link EmployeeServiceReadRequestDTO }
     * 
     * 
     */
    public List<EmployeeServiceReadRequestDTO> getEmployeeServiceReadRequestDTO() {
        if (employeeServiceReadRequestDTO == null) {
            employeeServiceReadRequestDTO = new ArrayList<EmployeeServiceReadRequestDTO>();
        }
        return this.employeeServiceReadRequestDTO;
    }

}
