
package com.mincom.enterpriseservice.screen;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;


/**
 * <p>Java class for ScreenDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="ScreenDTO">
 *   &lt;complexContent>
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
 *       &lt;sequence>
 *         &lt;element name="currentCursorFieldName" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *         &lt;element name="functionKeys" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *         &lt;element name="idle" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="mapFormUpdateDate" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *         &lt;element name="mapFormVersion" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *         &lt;element name="mapName" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *         &lt;element name="menu" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="message" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *         &lt;element name="nextAction" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *         &lt;element name="programName" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *         &lt;element name="screenFields" type="{http://screen.enterpriseservice.mincom.com}ArrayOfScreenFieldDTO" minOccurs="0"/>
 *         &lt;element name="screenTitle" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/restriction>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "ScreenDTO", propOrder = {
    "currentCursorFieldName",
    "functionKeys",
    "idle",
    "mapFormUpdateDate",
    "mapFormVersion",
    "mapName",
    "menu",
    "message",
    "nextAction",
    "programName",
    "screenFields",
    "screenTitle"
})
public class ScreenDTO {

    protected String currentCursorFieldName;
    protected String functionKeys;
    protected Boolean idle;
    protected String mapFormUpdateDate;
    protected String mapFormVersion;
    protected String mapName;
    protected Boolean menu;
    protected String message;
    protected String nextAction;
    protected String programName;
    protected ArrayOfScreenFieldDTO screenFields;
    protected String screenTitle;

    /**
     * Gets the value of the currentCursorFieldName property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getCurrentCursorFieldName() {
        return currentCursorFieldName;
    }

    /**
     * Sets the value of the currentCursorFieldName property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setCurrentCursorFieldName(String value) {
        this.currentCursorFieldName = value;
    }

    /**
     * Gets the value of the functionKeys property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getFunctionKeys() {
        return functionKeys;
    }

    /**
     * Sets the value of the functionKeys property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setFunctionKeys(String value) {
        this.functionKeys = value;
    }

    /**
     * Gets the value of the idle property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isIdle() {
        return idle;
    }

    /**
     * Sets the value of the idle property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setIdle(Boolean value) {
        this.idle = value;
    }

    /**
     * Gets the value of the mapFormUpdateDate property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getMapFormUpdateDate() {
        return mapFormUpdateDate;
    }

    /**
     * Sets the value of the mapFormUpdateDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setMapFormUpdateDate(String value) {
        this.mapFormUpdateDate = value;
    }

    /**
     * Gets the value of the mapFormVersion property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getMapFormVersion() {
        return mapFormVersion;
    }

    /**
     * Sets the value of the mapFormVersion property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setMapFormVersion(String value) {
        this.mapFormVersion = value;
    }

    /**
     * Gets the value of the mapName property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getMapName() {
        return mapName;
    }

    /**
     * Sets the value of the mapName property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setMapName(String value) {
        this.mapName = value;
    }

    /**
     * Gets the value of the menu property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isMenu() {
        return menu;
    }

    /**
     * Sets the value of the menu property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setMenu(Boolean value) {
        this.menu = value;
    }

    /**
     * Gets the value of the message property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getMessage() {
        return message;
    }

    /**
     * Sets the value of the message property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setMessage(String value) {
        this.message = value;
    }

    /**
     * Gets the value of the nextAction property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getNextAction() {
        return nextAction;
    }

    /**
     * Sets the value of the nextAction property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setNextAction(String value) {
        this.nextAction = value;
    }

    /**
     * Gets the value of the programName property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getProgramName() {
        return programName;
    }

    /**
     * Sets the value of the programName property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setProgramName(String value) {
        this.programName = value;
    }

    /**
     * Gets the value of the screenFields property.
     * 
     * @return
     *     possible object is
     *     {@link ArrayOfScreenFieldDTO }
     *     
     */
    public ArrayOfScreenFieldDTO getScreenFields() {
        return screenFields;
    }

    /**
     * Sets the value of the screenFields property.
     * 
     * @param value
     *     allowed object is
     *     {@link ArrayOfScreenFieldDTO }
     *     
     */
    public void setScreenFields(ArrayOfScreenFieldDTO value) {
        this.screenFields = value;
    }

    /**
     * Gets the value of the screenTitle property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getScreenTitle() {
        return screenTitle;
    }

    /**
     * Sets the value of the screenTitle property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setScreenTitle(String value) {
        this.screenTitle = value;
    }

}
