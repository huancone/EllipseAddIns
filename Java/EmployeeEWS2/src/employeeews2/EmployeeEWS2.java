/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package employeeews2;

import EllipseWebServicesClient.ClientConversation;
import com.mincom.enterpriseservice.ellipse.employee.Employee;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeService;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceCreateReplyDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceCreateRequestDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceModifyReplyDTO;
import com.mincom.enterpriseservice.ellipse.employee.EmployeeServiceModifyRequestDTO;
import com.mincom.enterpriseservice.ellipse.employee.EnterpriseServiceOperationException;
import com.mincom.ews.service.connectivity.OperationContext;
import java.net.MalformedURLException;
import java.net.URL;
import javax.swing.JOptionPane;

/**
 *
 * @author Hugo Andres
 */
public class EmployeeEWS2 {

    /**
     * @param args the command line arguments
     * @throws java.net.MalformedURLException
     * @throws
     * com.mincom.enterpriseservice.ellipse.employee.EnterpriseServiceOperationException
     */
    public static void main(String[] args) throws MalformedURLException, EnterpriseServiceOperationException {
        EmployeeService proxy = null;
        Employee Servicio = null;

        proxy = new EmployeeService(new URL("http://ews-el8desa.lmnerp03.cerrejon.com/ews/services/EmployeeService?WSDL"));

        Servicio = proxy.getEmployeeServiceHttpPort();

        EmployeeServiceCreateRequestDTO CreateParams = new EmployeeServiceCreateRequestDTO();
        EmployeeServiceCreateReplyDTO CreateReply = null;

        EmployeeServiceModifyRequestDTO ModifyParams = new EmployeeServiceModifyRequestDTO();
        EmployeeServiceModifyReplyDTO ModifyReply = new EmployeeServiceModifyReplyDTO();

        OperationContext CreateOpContext = new OperationContext();
        OperationContext ModifyOpContext = new OperationContext();

        ClientConversation.authenticate("VENTYX", "");

        CreateOpContext.setDistrict("ICOR");
        CreateOpContext.setPosition("");
        CreateOpContext.setReturnWarnings(true);
        CreateOpContext.setMaxInstances(100);

        ModifyOpContext.setDistrict("ICOR");
        ModifyOpContext.setPosition("");
        ModifyOpContext.setReturnWarnings(true);
        ModifyOpContext.setMaxInstances(100);

        try {
            CreateParams.setPersonType("CORE");
            CreateParams.setEmployee("HMENDOZA");
            CreateParams.setFirstName("HUGO");
            CreateParams.setLastName("ANDRES");
            CreateParams.setTitle("Mr");
            CreateParams.setPreferredName("91527581");
            CreateParams.setEmailAddress("hugo.mendoza@cerrejoncoal.com");
            CreateParams.setPrinterName1("PYT");
            CreateParams.setUnionCode("1");
            CreateParams.setNotifyEDIMsgRecieved(false);

            CreateReply = Servicio.create(CreateOpContext, CreateParams);
            JOptionPane.showMessageDialog(null, "Nombre: " + CreateReply.getEmployee());

            ModifyParams.setEmployee("HMENDOZA");
            ModifyParams.setFirstName("HUAN");
            ModifyReply = Servicio.modify(ModifyOpContext, ModifyParams);
            JOptionPane.showMessageDialog(null, "Nombre: " + CreateReply.getEmployee());

        } catch (EnterpriseServiceOperationException e) {
            JOptionPane.showMessageDialog(null, e.getMessage());
        }
    }

}
