
package com.mincom.enterpriseservice.ellipse;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlSeeAlso;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.securityclass.SecurityClassServiceModifyAttributesReplyDTO;
import com.mincom.enterpriseservice.ellipse.securityclass.SecurityClassServiceModifyMethodsReplyDTO;
import com.mincom.enterpriseservice.ellipse.securityclass.SecurityClassServiceModifyReplyDTO;
import com.mincom.enterpriseservice.ellipse.securityclass.SecurityClassServiceReadReplyDTO;
import com.mincom.enterpriseservice.ellipse.securityclass.SecurityClassServiceRetrieveClassesReplyDTO;
import com.mincom.enterpriseservice.ellipse.securityclass.SecurityClassServiceRetrieveReplyDTO;


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
    SecurityClassServiceModifyAttributesReplyDTO.class,
    SecurityClassServiceModifyReplyDTO.class,
    SecurityClassServiceReadReplyDTO.class,
    SecurityClassServiceModifyMethodsReplyDTO.class,
    SecurityClassServiceRetrieveClassesReplyDTO.class,
    SecurityClassServiceRetrieveReplyDTO.class
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
