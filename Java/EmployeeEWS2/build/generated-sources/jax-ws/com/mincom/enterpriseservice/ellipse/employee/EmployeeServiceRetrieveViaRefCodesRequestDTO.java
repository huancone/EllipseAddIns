
package com.mincom.enterpriseservice.ellipse.employee;

import java.math.BigDecimal;
import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.AbstractDTO;


/**
 * <p>Java class for EmployeeServiceRetrieveViaRefCodesRequestDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="EmployeeServiceRetrieveViaRefCodesRequestDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractDTO">
 *       &lt;sequence>
 *         &lt;element name="entityType" type="{http://employee.ellipse.enterpriseservice.mincom.com}entityType" minOccurs="0"/>
 *         &lt;element name="keyType" type="{http://employee.ellipse.enterpriseservice.mincom.com}keyType" minOccurs="0"/>
 *         &lt;element name="refCode" type="{http://employee.ellipse.enterpriseservice.mincom.com}refCode" minOccurs="0"/>
 *         &lt;element name="refNo" type="{http://employee.ellipse.enterpriseservice.mincom.com}refNo" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/extension>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "EmployeeServiceRetrieveViaRefCodesRequestDTO", propOrder = {
    "entityType",
    "keyType",
    "refCode",
    "refNo"
})
public class EmployeeServiceRetrieveViaRefCodesRequestDTO
    extends AbstractDTO
{

    protected String entityType;
    protected String keyType;
    protected String refCode;
    protected BigDecimal refNo;

    /**
     * Gets the value of the entityType property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getEntityType() {
        return entityType;
    }

    /**
     * Sets the value of the entityType property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setEntityType(String value) {
        this.entityType = value;
    }

    /**
     * Gets the value of the keyType property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getKeyType() {
        return keyType;
    }

    /**
     * Sets the value of the keyType property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setKeyType(String value) {
        this.keyType = value;
    }

    /**
     * Gets the value of the refCode property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getRefCode() {
        return refCode;
    }

    /**
     * Sets the value of the refCode property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setRefCode(String value) {
        this.refCode = value;
    }

    /**
     * Gets the value of the refNo property.
     * 
     * @return
     *     possible object is
     *     {@link BigDecimal }
     *     
     */
    public BigDecimal getRefNo() {
        return refNo;
    }

    /**
     * Sets the value of the refNo property.
     * 
     * @param value
     *     allowed object is
     *     {@link BigDecimal }
     *     
     */
    public void setRefNo(BigDecimal value) {
        this.refNo = value;
    }

}
