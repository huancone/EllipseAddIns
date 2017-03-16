
package com.mincom.enterpriseservice.ellipse.securityclass;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.AbstractReplyCollectionDTO;


/**
 * <p>Java class for SecurityClassServiceModifyReplyCollectionDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="SecurityClassServiceModifyReplyCollectionDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractReplyCollectionDTO">
 *       &lt;sequence>
 *         &lt;element name="replyElements" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}ArrayOfSecurityClassServiceModifyReplyDTO" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/extension>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "SecurityClassServiceModifyReplyCollectionDTO", propOrder = {
    "replyElements"
})
public class SecurityClassServiceModifyReplyCollectionDTO
    extends AbstractReplyCollectionDTO
{

    protected ArrayOfSecurityClassServiceModifyReplyDTO replyElements;

    /**
     * Gets the value of the replyElements property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfSecurityClassServiceModifyReplyDTO }
     *     
     */
    public ArrayOfSecurityClassServiceModifyReplyDTO getReplyElements() {
        return replyElements;
    }

    /**
     * Sets the value of the replyElements property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfSecurityClassServiceModifyReplyDTO }
     *     
     */
    public void setReplyElements(ArrayOfSecurityClassServiceModifyReplyDTO value) {
        this.replyElements = value;
    }

}
