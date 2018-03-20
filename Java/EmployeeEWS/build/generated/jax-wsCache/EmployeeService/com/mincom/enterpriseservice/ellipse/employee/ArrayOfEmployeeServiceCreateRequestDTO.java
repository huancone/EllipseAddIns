
package com.mincom.enterpriseservice.ellipse.employee;

import java.util.ArrayList;
import java.util.List;
import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlType;


/**
 * <p>Java class for ArrayOfEmployeeServiceCreateRequestDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="ArrayOfEmployeeServiceCreateRequestDTO">
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *       &lt;sequence>
 *         &lt;element name="EmployeeServiceCreateRequestDTO" type="{http://employee.ellipse.enterpriseservice.mincom.com}EmployeeServiceCreateRequestDTO" maxOccurs="unbounded" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "ArrayOfEmployeeServiceCreateRequestDTO", propOrder = {
    "employeeServiceCreateRequestDTO"
})
public class ArrayOfEmployeeServiceCreateRequestDTO {

    @XmlElement(name = "EmployeeServiceCreateRequestDTO", nillable = true)
    protected List<EmployeeServiceCreateRequestDTO> employeeServiceCreateRequestDTO;

    /**
     * Gets the value of the employeeServiceCreateRequestDTO property.
     * 
     * <p>
     * This accessor method returns a reference to the live list,
     * not a snapshot. Therefore any modification you make to the
     * returned list will be present inside the JAXB object.
     * This is why there is not a <CODE>set</CODE> method for the employeeServiceCreateRequestDTO property.
     * 
     * <p>
     * For example, to add a new item, do as follows:
     * <pre>
     *    getEmployeeServiceCreateRequestDTO().add(newItem);
     * </pre>
     * 
     * 
     * <p>
     * Objects of the following type(s) are allowed in the list
     * {@link EmployeeServiceCreateRequestDTO }
     * 
     * 
     */
    public List<EmployeeServiceCreateRequestDTO> getEmployeeServiceCreateRequestDTO() {
        if (employeeServiceCreateRequestDTO == null) {
            employeeServiceCreateRequestDTO = new ArrayList<EmployeeServiceCreateRequestDTO>();
        }
        return this.employeeServiceCreateRequestDTO;
    }

}
