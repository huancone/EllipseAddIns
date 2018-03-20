
package com.mincom.enterpriseservice.ellipse.employee;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.AbstractReplyCollectionDTO;


/**
 * <p>Java class for EmployeeServiceRetrieveReplyCollectionDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="EmployeeServiceRetrieveReplyCollectionDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractReplyCollectionDTO">
 *       &lt;sequence>
 *         &lt;element name="replyElements" type="{http://employee.ellipse.enterpriseservice.mincom.com}ArrayOfEmployeeServiceRetrieveReplyDTO" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/extension>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "EmployeeServiceRetrieveReplyCollectionDTO", propOrder = {
    "replyElements"
})
public class EmployeeServiceRetrieveReplyCollectionDTO
    extends AbstractReplyCollectionDTO
{

    protected ArrayOfEmployeeServiceRetrieveReplyDTO replyElements;

    /**
     * Gets the value of the replyElements property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfEmployeeServiceRetrieveReplyDTO }
     *     
     */
    public ArrayOfEmployeeServiceRetrieveReplyDTO getReplyElements() {
        return replyElements;
    }

    /**
     * Sets the value of the replyElements property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfEmployeeServiceRetrieveReplyDTO }
     *     
     */
    public void setReplyElements(ArrayOfEmployeeServiceRetrieveReplyDTO value) {
        this.replyElements = value;
    }

}
