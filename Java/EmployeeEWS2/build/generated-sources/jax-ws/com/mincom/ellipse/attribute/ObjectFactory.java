
package com.mincom.ellipse.attribute;

import javax.xml.bind.annotation.XmlRegistry;


/**
 * This object contains factory methods for each 
 * Java content interface and Java element interface 
 * generated in the com.mincom.ellipse.attribute package. 
 * <p>An ObjectFactory allows you to programatically 
 * construct new instances of the Java representation 
 * for XML content. The Java representation of XML 
 * content can consist of schema derived interfaces 
 * and classes representing the binding of schema 
 * type definitions, element declarations and model 
 * groups.  Factory methods for each of these are 
 * provided in this class.
 * 
 */
@XmlRegistry
public class ObjectFactory {


    /**
     * Create a new ObjectFactory that can be used to create new instances of schema derived classes for package: com.mincom.ellipse.attribute
     * 
     */
    public ObjectFactory() {
    }

    /**
     * Create an instance of {@link ArrayOfAttribute }
     * 
     */
    public ArrayOfAttribute createArrayOfAttribute() {
        return new ArrayOfAttribute();
    }

    /**
     * Create an instance of {@link Attribute }
     * 
     */
    public Attribute createAttribute() {
        return new Attribute();
    }

}
