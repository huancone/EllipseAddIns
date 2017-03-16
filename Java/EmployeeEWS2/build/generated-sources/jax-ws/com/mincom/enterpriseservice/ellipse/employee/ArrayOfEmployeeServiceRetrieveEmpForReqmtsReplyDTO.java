
package com.mincom.enterpriseservice.ellipse.employee;

import java.util.ArrayList;
import java.util.List;
import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlType;


/**
 * <p>Java class for ArrayOfEmployeeServiceRetrieveEmpForReqmtsReplyDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="ArrayOfEmployeeServiceRetrieveEmpForReqmtsReplyDTO">
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *       &lt;sequence>
 *         &lt;element name="EmployeeServiceRetrieveEmpForReqmtsReplyDTO" type="{http://employee.ellipse.enterpriseservice.mincom.com}EmployeeServiceRetrieveEmpForReqmtsReplyDTO" maxOccurs="unbounded" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "ArrayOfEmployeeServiceRetrieveEmpForReqmtsReplyDTO", propOrder = {
    "employeeServiceRetrieveEmpForReqmtsReplyDTO"
})
public class ArrayOfEmployeeServiceRetrieveEmpForReqmtsReplyDTO {

    @XmlElement(name = "EmployeeServiceRetrieveEmpForReqmtsReplyDTO", nillable = true)
    protected List<EmployeeServiceRetrieveEmpForReqmtsReplyDTO> employeeServiceRetrieveEmpForReqmtsReplyDTO;

    /**
     * Gets the value of the employeeServiceRetrieveEmpForReqmtsReplyDTO property.
     * 
     * <p>
     * This accessor method returns a reference to the live list,
     * not a snapshot. Therefore any modification you make to the
     * returned list will be present inside the JAXB object.
     * This is why there is not a <CODE>set</CODE> method for the employeeServiceRetrieveEmpForReqmtsReplyDTO property.
     * 
     * <p>
     * For example, to add a new item, do as follows:
     * <pre>
     *    getEmployeeServiceRetrieveEmpForReqmtsReplyDTO().add(newItem);
     * </pre>
     * 
     * 
     * <p>
     * Objects of the following type(s) are allowed in the list
     * {@link EmployeeServiceRetrieveEmpForReqmtsReplyDTO }
     * 
     * 
     */
    public List<EmployeeServiceRetrieveEmpForReqmtsReplyDTO> getEmployeeServiceRetrieveEmpForReqmtsReplyDTO() {
        if (employeeServiceRetrieveEmpForReqmtsReplyDTO == null) {
            employeeServiceRetrieveEmpForReqmtsReplyDTO = new ArrayList<EmployeeServiceRetrieveEmpForReqmtsReplyDTO>();
        }
        return this.employeeServiceRetrieveEmpForReqmtsReplyDTO;
    }

}
