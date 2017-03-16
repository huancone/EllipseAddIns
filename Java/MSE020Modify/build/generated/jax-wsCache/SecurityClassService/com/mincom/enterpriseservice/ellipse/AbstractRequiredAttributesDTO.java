
package com.mincom.enterpriseservice.ellipse;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlSeeAlso;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.securityclass.SecurityClassServiceModifyAttributesRequiredAttributesDTO;
import com.mincom.enterpriseservice.ellipse.securityclass.SecurityClassServiceModifyMethodsRequiredAttributesDTO;
import com.mincom.enterpriseservice.ellipse.securityclass.SecurityClassServiceModifyRequiredAttributesDTO;
import com.mincom.enterpriseservice.ellipse.securityclass.SecurityClassServiceReadRequiredAttributesDTO;
import com.mincom.enterpriseservice.ellipse.securityclass.SecurityClassServiceRetrieveClassesRequiredAttributesDTO;
import com.mincom.enterpriseservice.ellipse.securityclass.SecurityClassServiceRetrieveRequiredAttributesDTO;


/**
 * <p>Java class for AbstractRequiredAttributesDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="AbstractRequiredAttributesDTO">
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "AbstractRequiredAttributesDTO")
@XmlSeeAlso({
    SecurityClassServiceRetrieveRequiredAttributesDTO.class,
    SecurityClassServiceRetrieveClassesRequiredAttributesDTO.class,
    SecurityClassServiceModifyMethodsRequiredAttributesDTO.class,
    SecurityClassServiceModifyAttributesRequiredAttributesDTO.class,
    SecurityClassServiceReadRequiredAttributesDTO.class,
    SecurityClassServiceModifyRequiredAttributesDTO.class
})
public class AbstractRequiredAttributesDTO {


}
