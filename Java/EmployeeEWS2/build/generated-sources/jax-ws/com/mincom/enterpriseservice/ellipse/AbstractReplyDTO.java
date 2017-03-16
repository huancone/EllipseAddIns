
package com.mincom.enterpriseservice.ellipse;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlSeeAlso;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceCreateReplyDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceDeleteReplyDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceFetchLeaveHeaderReplyDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceModifyReplyDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceReadReplyDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceRetrieveEmpForReqmtsReplyDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceRetrieveForExtractReplyDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceRetrieveReplyDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceRetrieveViaRefCodesReplyDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceShowReplyDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceTransferPositionReplyDTO;


/**
 * <p>Java class for AbstractReplyDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="AbstractReplyDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractDTO">
 *       &lt;sequence>
 *         &lt;element name="warningsAndInformation" type="{http://ellipse.enterpriseservice.mincom.com}ArrayOfWarningMessageDTO" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/extension>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "AbstractReplyDTO", propOrder = {
    "warningsAndInformation"
})
@XmlSeeAlso({
    EmployeeServiceFetchLeaveHeaderReplyDTO.class,
    EmployeeServiceShowReplyDTO.class,
    EmployeeServiceCreateReplyDTO.class,
    EmployeeServiceReadReplyDTO.class,
    EmployeeServiceDeleteReplyDTO.class,
    EmployeeServiceModifyReplyDTO.class,
    EmployeeServiceTransferPositionReplyDTO.class,
    EmployeeServiceRetrieveReplyDTO.class,
    EmployeeServiceRetrieveEmpForReqmtsReplyDTO.class,
    EmployeeServiceRetrieveViaRefCodesReplyDTO.class,
    EmployeeServiceRetrieveForExtractReplyDTO.class
})
public class AbstractReplyDTO
    extends AbstractDTO
{

    protected ArrayOfWarningMessageDTO warningsAndInformation;

    /**
     * Gets the value of the warningsAndInformation property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfWarningMessageDTO }
     *     
     */
    public ArrayOfWarningMessageDTO getWarningsAndInformation() {
        return warningsAndInformation;
    }

    /**
     * Sets the value of the warningsAndInformation property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfWarningMessageDTO }
     *     
     */
    public void setWarningsAndInformation(ArrayOfWarningMessageDTO value) {
        this.warningsAndInformation = value;
    }

}
