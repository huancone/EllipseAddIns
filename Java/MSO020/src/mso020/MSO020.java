/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package mso020;

import EllipseWebServicesClient.ClientConversation;
import com.mincom.enterpriseservice.screen.ArrayOfScreenNameValueDTO;
import com.mincom.enterpriseservice.screen.EnterpriseServiceException;
import com.mincom.enterpriseservice.screen.InvalidConnectionIdException;
import com.mincom.enterpriseservice.screen.Screen;
import com.mincom.enterpriseservice.screen.ScreenDTO;
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
public class MSO020 {

    /**
     * @param args the command line arguments
     * @throws java.net.MalformedURLException
     * @throws com.mincom.enterpriseservice.screen.EnterpriseServiceException
     * @throws com.mincom.enterpriseservice.screen.InvalidConnectionIdException
     */
    public static void main(String[] args) throws MalformedURLException, EnterpriseServiceException, InvalidConnectionIdException {

        ScreenService proxy = null;
        ScreenDTO reply = new ScreenDTO();

        Screen servicio = null;

        ScreenSubmitRequestDTO request = new ScreenSubmitRequestDTO();

        proxy = new ScreenService(new URL("http://ews-el8prod.lmnerp01.cerrejon.com/ews/services/ScreenService?WSDL"));
        servicio = proxy.getScreenServiceHttpPort();

        OperationContext OpContext = new OperationContext();

        ClientConversation.authenticate("hmendo4", "abril06");

        OpContext.setDistrict("ICOR");
        OpContext.setPosition("TOP");
        OpContext.setReturnWarnings(true);
        OpContext.setMaxInstances(100);

        try {
            reply = servicio.executeScreen(OpContext, "MSO020");
        } catch (EnterpriseServiceException | InvalidConnectionIdException e) {
            JOptionPane.showMessageDialog(null, e.getMessage());
        }

        request.setScreenFields(null);
        request.setScreenKey("3");

        try {
            reply = servicio.submit(OpContext, request);
            reply = servicio.executeScreen(OpContext, "MSO020");
        } catch (EnterpriseServiceException | InvalidConnectionIdException e) {
            JOptionPane.showMessageDialog(null, e.getMessage());
        }

        if (reply.getMapName().equals("MSM020A")) {

            request = new ScreenSubmitRequestDTO();
            ArrayOfScreenNameValueDTO values = new ArrayOfScreenNameValueDTO();

            ScreenNameValueDTO field1 = new ScreenNameValueDTO();
            ScreenNameValueDTO field2 = new ScreenNameValueDTO();
            ScreenNameValueDTO field3 = new ScreenNameValueDTO();
            ScreenNameValueDTO field4 = new ScreenNameValueDTO();
            ScreenNameValueDTO field5 = new ScreenNameValueDTO();

            field1.setFieldName("OPTION1I");
            field1.setValue("2");

            field2.setFieldName("ENTRY_TYPE1I");
            field2.setValue("S");

            field3.setFieldName("ENTITY_A1I");
            field3.setValue("COMC2");

            field4.setFieldName("FORMAT1I");
            field4.setValue("S");

            field5.setFieldName("DSTRCT_CODE_A1I");
            field5.setValue("ICOR");

            values.getScreenNameValueDTO().add(field1);
            values.getScreenNameValueDTO().add(field2);
            values.getScreenNameValueDTO().add(field3);
            values.getScreenNameValueDTO().add(field4);
            values.getScreenNameValueDTO().add(field5);

            request.setScreenFields(values);
            request.setScreenKey("1");
        }

        try {
            reply = servicio.submit(OpContext, request);
        } catch (EnterpriseServiceException | InvalidConnectionIdException e) {
            JOptionPane.showMessageDialog(null, e.getMessage());
        }

        if (reply.getMapName().equals("MSM020C")) {
            request = new ScreenSubmitRequestDTO();
            ArrayOfScreenNameValueDTO values = new ArrayOfScreenNameValueDTO();

            ScreenNameValueDTO field1 = new ScreenNameValueDTO();
            ScreenNameValueDTO field2 = new ScreenNameValueDTO();
            ScreenNameValueDTO field3 = new ScreenNameValueDTO();

            field1.setFieldName("DEFAULT_MENU3I");
            field1.setValue("GM");
            
            field2.setFieldName("DEF_FOR_EMP3I");
            field2.setValue("N");
            
            field3.setFieldName("PRF_LOGIN_LCKD3I");
            field3.setValue("N");
            
            values.getScreenNameValueDTO().add(field1);
            values.getScreenNameValueDTO().add(field2);
            values.getScreenNameValueDTO().add(field3);

            request.setScreenFields(values);
            request.setScreenKey("1");
        }

        try {
            reply = servicio.submit(OpContext, request);
            JOptionPane.showMessageDialog(null, "Creado");
        } catch (EnterpriseServiceException | InvalidConnectionIdException e) {
            JOptionPane.showMessageDialog(null, e.getMessage());
        }

        request.setScreenFields(null);
        request.setScreenKey("3");

        reply = servicio.submit(OpContext, request);
    }

}
