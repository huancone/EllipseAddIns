
package com.mincom.enterpriseservice.ellipse.employee;

import java.util.ArrayList;
import java.util.List;
import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlType;


/**
 * <p>Java class for ArrayOfEmployeeServiceRetrieveForExtractReplyDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="ArrayOfEmployeeServiceRetrieveForExtractReplyDTO">
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *       &lt;sequence>
 *         &lt;element name="EmployeeServiceRetrieveForExtractReplyDTO" type="{http://employee.ellipse.enterpriseservice.mincom.com}EmployeeServiceRetrieveForExtractReplyDTO" maxOccurs="unbounded" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "ArrayOfEmployeeServiceRetrieveForExtractReplyDTO", propOrder = {
    "employeeServiceRetrieveForExtractReplyDTO"
})
public class ArrayOfEmployeeServiceRetrieveForExtractReplyDTO {

    @XmlElement(name = "EmployeeServiceRetrieveForExtractReplyDTO", nillable = true)
    protected List<EmployeeServiceRetrieveForExtractReplyDTO> employeeServiceRetrieveForExtractReplyDTO;

    /**
     * Gets the value of the employeeServiceRetrieveForExtractReplyDTO property.
     * 
     * <p>
     * This accessor method returns a reference to the live list,
     * not a snapshot. Therefore any modification you make to the
     * returned list will be present inside the JAXB object.
     * This is why there is not a <CODE>set</CODE> method for the employeeServiceRetrieveForExtractReplyDTO property.
     * 
     * <p>
     * For example, to add a new item, do as follows:
     * <pre>
     *    getEmployeeServiceRetrieveForExtractReplyDTO().add(newItem);
     * </pre>
     * 
     * 
     * <p>
     * Objects of the following type(s) are allowed in the list
     * {@link EmployeeServiceRetrieveForExtractReplyDTO }
     * 
     * 
     */
    public List<EmployeeServiceRetrieveForExtractReplyDTO> getEmployeeServiceRetrieveForExtractReplyDTO() {
        if (employeeServiceRetrieveForExtractReplyDTO == null) {
            employeeServiceRetrieveForExtractReplyDTO = new ArrayList<EmployeeServiceRetrieveForExtractReplyDTO>();
        }
        return this.employeeServiceRetrieveForExtractReplyDTO;
    }

}
