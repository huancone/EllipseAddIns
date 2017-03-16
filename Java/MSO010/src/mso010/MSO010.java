/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package mso010;

import EllipseWebServicesClient.ClientConversation;
import com.mincom.enterpriseservice.screen.ArrayOfScreenFieldDTO;
import com.mincom.enterpriseservice.screen.ArrayOfScreenNameValueDTO;
import com.mincom.enterpriseservice.screen.EnterpriseServiceException;
import com.mincom.enterpriseservice.screen.InvalidConnectionIdException;
import com.mincom.enterpriseservice.screen.Screen;
import com.mincom.enterpriseservice.screen.ScreenDTO;
import com.mincom.enterpriseservice.screen.ScreenFieldDTO;
import com.mincom.enterpriseservice.screen.ScreenNameValueDTO;
import com.mincom.enterpriseservice.screen.ScreenService;
import com.mincom.enterpriseservice.screen.ScreenSubmitRequestDTO;
import com.mincom.ews.service.connectivity.OperationContext;
import java.net.MalformedURLException;
import java.net.URL;
import javax.swing.JOptionPane;

/**
 *
 * @author Hugo Andres
 */
public class MSO010 {

    /**
     * @param args the command line arguments
     * @throws java.net.MalformedURLException
     */
    public static void main(String[] args) throws MalformedURLException, EnterpriseServiceException, InvalidConnectionIdException {

        ScreenService proxy = null;
        ScreenDTO reply = new ScreenDTO();
        Screen servicio = null;
        ScreenSubmitRequestDTO request = new ScreenSubmitRequestDTO();
        proxy = new ScreenService(new URL("http://ews-el8desa.lmnerp03.cerrejon.com/ews/services/ScreenService?WSDL"));
        servicio = proxy.getScreenServiceHttpPort();

        OperationContext OpContext = new OperationContext();

        ClientConversation.authenticate("EROMERO", "");

        OpContext.setDistrict("ICOR");
        OpContext.setPosition("TOP");
        OpContext.setReturnWarnings(true);
        OpContext.setMaxInstances(100);
        
        try {
            reply = servicio.executeScreen(OpContext, "MSO010");
            JOptionPane.showMessageDialog(null, reply.getMapName());
        } catch (EnterpriseServiceException | InvalidConnectionIdException e) {
            JOptionPane.showMessageDialog(null, e.getMessage());
        }

        request.setScreenFields(null);
        request.setScreenKey("3");

        try {
            reply = servicio.submit(OpContext, request);
            reply = servicio.executeScreen(OpContext, "MSO010");
        } catch (EnterpriseServiceException | InvalidConnectionIdException e) {
            JOptionPane.showMessageDialog(null, e.getMessage());
        }

        if (reply.getMapName().equals("MSM010A")) {

            request = new ScreenSubmitRequestDTO();
            ArrayOfScreenNameValueDTO values = new ArrayOfScreenNameValueDTO();

            ScreenNameValueDTO field = new ScreenNameValueDTO();
            ScreenNameValueDTO field2 = new ScreenNameValueDTO();

            field.setFieldName("OPTION1I");
            field.setValue("1");

            field2.setFieldName("TABLE_TYPE1I");
            field2.setValue("WO");

            values.getScreenNameValueDTO().add(field);
            values.getScreenNameValueDTO().add(field2);

            request.setScreenFields(values);
            request.setScreenKey("1");
        }

        try {
            reply = servicio.submit(OpContext, request);
        } catch (EnterpriseServiceException | InvalidConnectionIdException e) {
            JOptionPane.showMessageDialog(null, e.getMessage());
        }

        if (reply.getMapName().equals("MSM010B")) {
            request = new ScreenSubmitRequestDTO();
            ArrayOfScreenNameValueDTO values = new ArrayOfScreenNameValueDTO();

            ScreenNameValueDTO field = new ScreenNameValueDTO();
            ScreenNameValueDTO field2 = new ScreenNameValueDTO();
            ScreenNameValueDTO field3 = new ScreenNameValueDTO();

            field.setFieldName("TABLE_CODE_A2I1");
            field.setValue("TE");

            field2.setFieldName("TABLE_DESC2I1");
            field2.setValue("TAREAS ELLIPSE");

            field3.setFieldName("ACTIVE_FLAG2I1");
            field3.setValue("Y");

            values.getScreenNameValueDTO().add(field);
            values.getScreenNameValueDTO().add(field2);
            values.getScreenNameValueDTO().add(field3);

            request.setScreenFields(values);
            request.setScreenKey("1");
        }

        try {
            reply = servicio.submit(OpContext, request);
            JOptionPane.showMessageDialog(null, reply.getMessage());
        } catch (EnterpriseServiceException | InvalidConnectionIdException e) {
            JOptionPane.showMessageDialog(null, e.getMessage());
        }

        request.setScreenFields(null);
        request.setScreenKey("3");

        reply = servicio.submit(OpContext, request);
    }

}
