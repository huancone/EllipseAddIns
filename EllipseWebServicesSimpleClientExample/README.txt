SimpleEWSClient software

THIS SOFTWARE IS PROVIDED WITHOUT A WARRANTY, WHETHER EXPLICIT OR IMPLIED.

To build the SimpleEWSClient program
====================================

Requires JDK 1.6 or greater.

From the nexus repository get the appropriate EllipseWebServicesJavaClient.zip artifact, extract the EWSJavaClient.jar and
place in this projects' lib folder.

Alternatively, locate the EWS application deployed to the application server, make a copy of the EllipseWebServicesJavaClient.zip artefact,
extract the EWSJavaClient.jar and place in this projects' lib folder. 

Alternatively use a web browser and download the EllipseWebServicesJavaClient.zip artifact from the link appearing on the
Ellipse Web Services information page (http:|https://<host name|ip address>[:port]/ews/). Repeat the remaining procedure as 
described above.

To run the SimpleEWSClient program
==================================

Requires JVM 1.6 or greater.

Java runtime class path

Ensure the project META-INF/cxf.xml is on the class path.

The following java archive files (?.jar) need to be on the class path.

EWSJavaClient.jar (Ellipse baseline specific)

commons-lang (v2.6)
commons-logging (v1.1.1)

spring framework (v3.0.6.RELEASE)

spring-beans
spring-context
spring-core
spring-expression
spring-asm
spring-aop


Crossfire (cxf.apache.org)

cxf.jar  (v2.5.2)

Crossfire/JAX-WS dependencies

neethi (v3.0.1)
wsdl4j (v1.6.2)
wss4j (v1.6.4)
xmlschema-core (v2.0.1)
xmlsec (v1.4.6)

Tested with versions of these components as denoted by values in parenthesis.  Earlier versions may be problematic. 


Command line parameters:

user                   Ellipse user
password  (optional)   Ellipse user password , do not provide this parameter if the Ellipse user has no password.
position               Ellipse user profile position
district               Ellipse user profile district
host                   Server hosting the EWS services , or the load balancing server used in front of the EWS server
port                   Tcp port EWS server (or load balancer) is listening for EWS client SOAP requests.


provide name, value pair combination for each parameter, e.g.

user=am2122 password=apassword position=SYSAD district=R100 host=ewshost port=8080

