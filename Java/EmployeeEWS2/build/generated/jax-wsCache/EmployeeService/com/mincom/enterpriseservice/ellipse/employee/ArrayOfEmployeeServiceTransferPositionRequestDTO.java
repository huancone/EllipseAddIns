
package com.mincom.enterpriseservice.ellipse.employee;

import java.util.ArrayList;
import java.util.List;
import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlType;


/**
 * <p>Java class for ArrayOfEmployeeServiceTransferPositionRequestDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="ArrayOfEmployeeServiceTransferPositionRequestDTO">
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *       &lt;sequence>
 *         &lt;element name="EmployeeServiceTransferPositionRequestDTO" type="{http://employee.ellipse.enterpriseservice.mincom.com}EmployeeServiceTransferPositionRequestDTO" maxOccurs="unbounded" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "ArrayOfEmployeeServiceTransferPositionRequestDTO", propOrder = {
    "employeeServiceTransferPositionRequestDTO"
})
public class ArrayOfEmployeeServiceTransferPositionRequestDTO {

    @XmlElement(name = "EmployeeServiceTransferPositionRequestDTO", nillable = true)
    protected List<EmployeeServiceTransferPositionRequestDTO> employeeServiceTransferPositionRequestDTO;

    /**
     * Gets the value of the employeeServiceTransferPositionRequestDTO property.
     * 
     * <p>
     * This accessor method returns a reference to the live list,
     * not a snapshot. Therefore any modification you make to the
     * returned list will be present inside the JAXB object.
     * This is why there is not a <CODE>set</CODE> method for the employeeServiceTransferPositionRequestDTO property.
     * 
     * <p>
     * For example, to add a new item, do as follows:
     * <pre>
     *    getEmployeeServiceTransferPositionRequestDTO().add(newItem);
     * </pre>
     * 
     * 
     * <p>
     * Objects of the following type(s) are allowed in the list
     * {@link EmployeeServiceTransferPositionRequestDTO }
     * 
     * 
     */
    public List<EmployeeServiceTransferPositionRequestDTO> getEmployeeServiceTransferPositionRequestDTO() {
        if (employeeServiceTransferPositionRequestDTO == null) {
            employeeServiceTransferPositionRequestDTO = new ArrayList<EmployeeServiceTransferPositionRequestDTO>();
        }
        return this.employeeServiceTransferPositionRequestDTO;
    }

}
