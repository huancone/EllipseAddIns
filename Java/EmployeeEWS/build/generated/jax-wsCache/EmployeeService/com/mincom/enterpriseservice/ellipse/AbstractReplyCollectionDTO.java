
package com.mincom.enterpriseservice.ellipse;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlSeeAlso;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceCreateReplyCollectionDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceDeleteReplyCollectionDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceFetchLeaveHeaderReplyCollectionDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceModifyReplyCollectionDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceReadReplyCollectionDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceRetrieveEmpForReqmtsReplyCollectionDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceRetrieveForExtractReplyCollectionDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceRetrieveReplyCollectionDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceRetrieveViaRefCodesReplyCollectionDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceShowReplyCollectionDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceTransferPositionReplyCollectionDTO;


/**
 * <p>Java class for AbstractReplyCollectionDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="AbstractReplyCollectionDTO">
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *       &lt;sequence>
 *         &lt;element name="collectionRestartPoint" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "AbstractReplyCollectionDTO", propOrder = {
    "collectionRestartPoint"
})
@XmlSeeAlso({
    EmployeeServiceReadReplyCollectionDTO.class,
    EmployeeServiceDeleteReplyCollectionDTO.class,
    EmployeeServiceFetchLeaveHeaderReplyCollectionDTO.class,
    EmployeeServiceCreateReplyCollectionDTO.class,
    EmployeeServiceRetrieveForExtractReplyCollectionDTO.class,
    EmployeeServiceRetrieveReplyCollectionDTO.class,
    EmployeeServiceModifyReplyCollectionDTO.class,
    EmployeeServiceRetrieveViaRefCodesReplyCollectionDTO.class,
    EmployeeServiceRetrieveEmpForReqmtsReplyCollectionDTO.class,
    EmployeeServiceTransferPositionReplyCollectionDTO.class,
    EmployeeServiceShowReplyCollectionDTO.class
})
public class AbstractReplyCollectionDTO {

    protected String collectionRestartPoint;

    /**
     * Gets the value of the collectionRestartPoint property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getCollectionRestartPoint() {
        return collectionRestartPoint;
    }

    /**
     * Sets the value of the collectionRestartPoint property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setCollectionRestartPoint(String value) {
        this.collectionRestartPoint = value;
    }

}
