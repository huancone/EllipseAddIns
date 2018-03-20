
package com.mincom.enterpriseservice.ellipse;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlSeeAlso;
import javax.xml.bind.annotation.XmlType;
import com.mincom.ellipse.attribute.ArrayOfAttribute;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceCreateRequestDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceDeleteRequestDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceFetchLeaveHeaderRequestDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceModifyRequestDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceReadRequestDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceRetrieveEmpForReqmtsRequestDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceRetrieveForExtractRequestDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceRetrieveRequestDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceRetrieveViaRefCodesRequestDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceShowRequestDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceTransferPositionRequestDTO;


/**
 * <p>Java class for AbstractDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="AbstractDTO">
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *       &lt;sequence>
 *         &lt;element name="customAttributes" type="{http://attribute.ellipse.mincom.com}ArrayOfAttribute" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "AbstractDTO", propOrder = {
    "customAttributes"
})
@XmlSeeAlso({
    EmployeeServiceShowRequestDTO.class,
    EmployeeServiceRetrieveViaRefCodesRequestDTO.class,
    EmployeeServiceRetrieveForExtractRequestDTO.class,
    EmployeeServiceTransferPositionRequestDTO.class,
    EmployeeServiceDeleteRequestDTO.class,
    EmployeeServiceReadRequestDTO.class,
    EmployeeServiceRetrieveEmpForReqmtsRequestDTO.class,
    EmployeeServiceModifyRequestDTO.class,
    EmployeeServiceFetchLeaveHeaderRequestDTO.class,
    EmployeeServiceCreateRequestDTO.class,
    EmployeeServiceRetrieveRequestDTO.class,
    AbstractReplyDTO.class
})
public class AbstractDTO {

    protected ArrayOfAttribute customAttributes;

    /**
     * Gets the value of the customAttributes property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfAttribute }
     *     
     */
    public ArrayOfAttribute getCustomAttributes() {
        return customAttributes;
    }

    /**
     * Sets the value of the customAttributes property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfAttribute }
     *     
     */
    public void setCustomAttributes(ArrayOfAttribute value) {
        this.customAttributes = value;
    }

}
