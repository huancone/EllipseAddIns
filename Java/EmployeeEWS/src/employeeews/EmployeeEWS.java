/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package employeeews;

import EllipseWebServicesClient.ClientConversation;
import com.mincom.enterpriseservice.ellipse.employee.Employee;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeService;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceReadReplyDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceReadRequestDTO;
import com.mincom.enterpriseservice.ellipse.employee.EnterpriseServiceOperationException;
import com.mincom.ews.service.connectivity.OperationContext;
import java.awt.HeadlessException;
import java.net.MalformedURLException;
import java.net.URL;
import javax.swing.JOptionPane;

/**
 *
 * @author Hugo Andres
 */
public class EmployeeEWS {
    private static OperationContext opc;
    /**
     * @param args the command line arguments
     * @throws java.net.MalformedURLException
     */
    public static void main(String[] args) throws MalformedURLException {
        EmployeeService proxy;
        proxy = null;
        Employee servicio = null;

        proxy = new EmployeeService(new URL("http://ews-el8prod.lmnerp02.cerrejon.com/ews/services/EmployeeService?WSDL"));
        servicio = proxy.getEmployeeServiceHttpPort();

        EmployeeServiceReadRequestDTO param = new EmployeeServiceReadRequestDTO();
        EmployeeServiceReadReplyDTO reply = null;
        OperationContext OpContext = new OperationContext();

        ClientConversation.authenticate("GGOMEZ", "");

        OpContext.setDistrict("ICOR");
        OpContext.setPosition("ADMIN");
        OpContext.setReturnWarnings(true);
        OpContext.setMaxInstances(100);

        param.setEmployee("GGOMEZ");

        try {
            reply = servicio.read(OpContext, param);
            JOptionPane.showMessageDialog(null, "Nombre: " + reply.getFirstName());
        } catch (EnterpriseServiceOperationException | HeadlessException e) {
            JOptionPane.showMessageDialog(null, e.getMessage());
        }
    }
}
