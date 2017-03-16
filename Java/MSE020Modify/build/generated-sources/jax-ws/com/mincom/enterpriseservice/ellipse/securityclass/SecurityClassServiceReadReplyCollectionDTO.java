
package com.mincom.enterpriseservice.ellipse.securityclass;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.AbstractReplyCollectionDTO;


/**
 * <p>Java class for SecurityClassServiceReadReplyCollectionDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="SecurityClassServiceReadReplyCollectionDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractReplyCollectionDTO">
 *       &lt;sequence>
 *         &lt;element name="replyElements" type="{http://securityclass.ellipse.enterpriseservice.mincom.com}ArrayOfSecurityClassServiceReadReplyDTO" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/extension>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "SecurityClassServiceReadReplyCollectionDTO", propOrder = {
    "replyElements"
})
public class SecurityClassServiceReadReplyCollectionDTO
    extends AbstractReplyCollectionDTO
{

    protected ArrayOfSecurityClassServiceReadReplyDTO replyElements;

    /**
     * Gets the value of the replyElements property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfSecurityClassServiceReadReplyDTO }
     *     
     */
    public ArrayOfSecurityClassServiceReadReplyDTO getReplyElements() {
        return replyElements;
    }

    /**
     * Sets the value of the replyElements property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfSecurityClassServiceReadReplyDTO }
     *     
     */
    public void setReplyElements(ArrayOfSecurityClassServiceReadReplyDTO value) {
        this.replyElements = value;
    }

}
