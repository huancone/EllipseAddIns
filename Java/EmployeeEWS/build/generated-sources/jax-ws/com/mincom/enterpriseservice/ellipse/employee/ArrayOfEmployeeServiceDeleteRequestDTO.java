
package com.mincom.enterpriseservice.ellipse.employee;

import java.util.ArrayList;
import java.util.List;
import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlType;


/**
 * <p>Java class for ArrayOfEmployeeServiceDeleteRequestDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="ArrayOfEmployeeServiceDeleteRequestDTO">
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *       &lt;sequence>
 *         &lt;element name="EmployeeServiceDeleteRequestDTO" type="{http://employee.ellipse.enterpriseservice.mincom.com}EmployeeServiceDeleteRequestDTO" maxOccurs="unbounded" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "ArrayOfEmployeeServiceDeleteRequestDTO", propOrder = {
    "employeeServiceDeleteRequestDTO"
})
public class ArrayOfEmployeeServiceDeleteRequestDTO {

    @XmlElement(name = "EmployeeServiceDeleteRequestDTO", nillable = true)
    protected List<EmployeeServiceDeleteRequestDTO> employeeServiceDeleteRequestDTO;

    /**
     * Gets the value of the employeeServiceDeleteRequestDTO property.
     * 
     * <p>
     * This accessor method returns a reference to the live list,
     * not a snapshot. Therefore any modification you make to the
     * returned list will be present inside the JAXB object.
     * This is why there is not a <CODE>set</CODE> method for the employeeServiceDeleteRequestDTO property.
     * 
     * <p>
     * For example, to add a new item, do as follows:
     * <pre>
     *    getEmployeeServiceDeleteRequestDTO().add(newItem);
     * </pre>
     * 
     * 
     * <p>
     * Objects of the following type(s) are allowed in the list
     * {@link EmployeeServiceDeleteRequestDTO }
     * 
     * 
     */
    public List<EmployeeServiceDeleteRequestDTO> getEmployeeServiceDeleteRequestDTO() {
        if (employeeServiceDeleteRequestDTO == null) {
            employeeServiceDeleteRequestDTO = new ArrayList<EmployeeServiceDeleteRequestDTO>();
        }
        return this.employeeServiceDeleteRequestDTO;
    }

}
