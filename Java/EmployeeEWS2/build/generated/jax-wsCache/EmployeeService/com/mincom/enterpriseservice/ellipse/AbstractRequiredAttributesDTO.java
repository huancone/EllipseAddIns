
package com.mincom.enterpriseservice.ellipse;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlSeeAlso;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceCreateRequiredAttributesDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceDeleteRequiredAttributesDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceFetchLeaveHeaderRequiredAttributesDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceModifyRequiredAttributesDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceReadRequiredAttributesDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceRetrieveEmpForReqmtsRequiredAttributesDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceRetrieveForExtractRequiredAttributesDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceRetrieveRequiredAttributesDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceRetrieveViaRefCodesRequiredAttributesDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceShowRequiredAttributesDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceTransferPositionRequiredAttributesDTO;


/**
 * <p>Java class for AbstractRequiredAttributesDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="AbstractRequiredAttributesDTO">
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "AbstractRequiredAttributesDTO")
@XmlSeeAlso({
    EmployeeServiceRetrieveViaRefCodesRequiredAttributesDTO.class,
    EmployeeServiceRetrieveForExtractRequiredAttributesDTO.class,
    EmployeeServiceRetrieveEmpForReqmtsRequiredAttributesDTO.class,
    EmployeeServiceRetrieveRequiredAttributesDTO.class,
    EmployeeServiceFetchLeaveHeaderRequiredAttributesDTO.class,
    EmployeeServiceModifyRequiredAttributesDTO.class,
    EmployeeServiceShowRequiredAttributesDTO.class,
    EmployeeServiceDeleteRequiredAttributesDTO.class,
    EmployeeServiceTransferPositionRequiredAttributesDTO.class,
    EmployeeServiceReadRequiredAttributesDTO.class,
    EmployeeServiceCreateRequiredAttributesDTO.class
})
public class AbstractRequiredAttributesDTO {


}
