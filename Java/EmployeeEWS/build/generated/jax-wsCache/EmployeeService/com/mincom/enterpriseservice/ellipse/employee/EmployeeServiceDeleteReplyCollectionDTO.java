
package com.mincom.enterpriseservice.ellipse.employee;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.AbstractReplyCollectionDTO;


/**
 * <p>Java class for EmployeeServiceDeleteReplyCollectionDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="EmployeeServiceDeleteReplyCollectionDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractReplyCollectionDTO">
 *       &lt;sequence>
 *         &lt;element name="replyElements" type="{http://employee.ellipse.enterpriseservice.mincom.com}ArrayOfEmployeeServiceDeleteReplyDTO" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/extension>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "EmployeeServiceDeleteReplyCollectionDTO", propOrder = {
    "replyElements"
})
public class EmployeeServiceDeleteReplyCollectionDTO
    extends AbstractReplyCollectionDTO
{

    protected ArrayOfEmployeeServiceDeleteReplyDTO replyElements;

    /**
     * Gets the value of the replyElements property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfEmployeeServiceDeleteReplyDTO }
     *     
     */
    public ArrayOfEmployeeServiceDeleteReplyDTO getReplyElements() {
        return replyElements;
    }

    /**
     * Sets the value of the replyElements property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfEmployeeServiceDeleteReplyDTO }
     *     
     */
    public void setReplyElements(ArrayOfEmployeeServiceDeleteReplyDTO value) {
        this.replyElements = value;
    }

}
