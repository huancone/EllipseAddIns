
package com.mincom.enterpriseservice.ellipse.employee;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.AbstractReplyCollectionDTO;


/**
 * <p>Java class for EmployeeServiceReadReplyCollectionDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="EmployeeServiceReadReplyCollectionDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractReplyCollectionDTO">
 *       &lt;sequence>
 *         &lt;element name="replyElements" type="{http://employee.ellipse.enterpriseservice.mincom.com}ArrayOfEmployeeServiceReadReplyDTO" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/extension>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "EmployeeServiceReadReplyCollectionDTO", propOrder = {
    "replyElements"
})
public class EmployeeServiceReadReplyCollectionDTO
    extends AbstractReplyCollectionDTO
{

    protected ArrayOfEmployeeServiceReadReplyDTO replyElements;

    /**
     * Gets the value of the replyElements property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfEmployeeServiceReadReplyDTO }
     *     
     */
    public ArrayOfEmployeeServiceReadReplyDTO getReplyElements() {
        return replyElements;
    }

    /**
     * Sets the value of the replyElements property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfEmployeeServiceReadReplyDTO }
     *     
     */
    public void setReplyElements(ArrayOfEmployeeServiceReadReplyDTO value) {
        this.replyElements = value;
    }

}
