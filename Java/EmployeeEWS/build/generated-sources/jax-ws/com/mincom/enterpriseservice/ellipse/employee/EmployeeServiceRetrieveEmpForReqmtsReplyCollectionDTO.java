
package com.mincom.enterpriseservice.ellipse.employee;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.AbstractReplyCollectionDTO;


/**
 * <p>Java class for EmployeeServiceRetrieveEmpForReqmtsReplyCollectionDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="EmployeeServiceRetrieveEmpForReqmtsReplyCollectionDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractReplyCollectionDTO">
 *       &lt;sequence>
 *         &lt;element name="replyElements" type="{http://employee.ellipse.enterpriseservice.mincom.com}ArrayOfEmployeeServiceRetrieveEmpForReqmtsReplyDTO" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/extension>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "EmployeeServiceRetrieveEmpForReqmtsReplyCollectionDTO", propOrder = {
    "replyElements"
})
public class EmployeeServiceRetrieveEmpForReqmtsReplyCollectionDTO
    extends AbstractReplyCollectionDTO
{

    protected ArrayOfEmployeeServiceRetrieveEmpForReqmtsReplyDTO replyElements;

    /**
     * Gets the value of the replyElements property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfEmployeeServiceRetrieveEmpForReqmtsReplyDTO }
     *     
     */
    public ArrayOfEmployeeServiceRetrieveEmpForReqmtsReplyDTO getReplyElements() {
        return replyElements;
    }

    /**
     * Sets the value of the replyElements property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfEmployeeServiceRetrieveEmpForReqmtsReplyDTO }
     *     
     */
    public void setReplyElements(ArrayOfEmployeeServiceRetrieveEmpForReqmtsReplyDTO value) {
        this.replyElements = value;
    }

}
