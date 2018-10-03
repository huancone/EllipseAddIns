package com.mincom.ews.client.example;

import java.net.URL;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.StringTokenizer;

import com.mincom.enterpriseservice.ellipse.ErrorMessageDTO;
import com.mincom.enterpriseservice.ellipse.equipment.EnterpriseServiceOperationException;
import com.mincom.enterpriseservice.ellipse.table.ArrayOfTableServiceRetrieveReplyDTO;
import com.mincom.enterpriseservice.ellipse.table.Table;
import com.mincom.enterpriseservice.ellipse.table.TableServiceRetrieveReplyCollectionDTO;
import com.mincom.enterpriseservice.ellipse.table.TableServiceRetrieveReplyDTO;
import com.mincom.enterpriseservice.ellipse.table.TableServiceRetrieveRequestDTO;
import com.mincom.ews.client.EWSClientConversation;
import com.mincom.ews.service.connectivity.OperationContext;

public class SimpleEWSClient {

	private Map<String,String> cmdLineParameters = new HashMap<String,String>();
	private static String[] cmdLineParameterNames = {"user", "password", "position", "district", "host", "port"};
	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		SimpleEWSClient executor = new SimpleEWSClient();
		try {
			executor.processCmdLineParameters(args);
			executor.callEWSService();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
	private void processCmdLineParameters(String[] args) throws Exception {
		for (int i=0;i<cmdLineParameterNames.length;i++) {
			cmdLineParameters.put(cmdLineParameterNames[i], "");
		}
		for (int i=0;i<args.length;i++) {
			StringTokenizer sTokenizer = new StringTokenizer(args[i],"=");
			String paramName = null;
			if (sTokenizer.countTokens()>0) {
				paramName = sTokenizer.nextToken().toLowerCase();
			} else {
				throw new Exception("error in comand line parameters");
			}
			if (sTokenizer.countTokens()>=1) {
				String value = sTokenizer.nextToken();
				String oldValue =cmdLineParameters.put(paramName, value);
				if (oldValue==null) {
					throw new Exception("Unknown parameter = "+paramName);
				}
			} else {
				throw new Exception("error in comand line parameters");
			}
		}
		String exceptionCause = "missing command line parameters: ";
		StringBuilder sBuilder = new StringBuilder(exceptionCause);
		for (String paramName :  cmdLineParameters.keySet()) {
			if ("".equals(cmdLineParameters.get(paramName))&& !"password".equals(paramName)) {
				sBuilder.append(paramName);
				sBuilder.append(", ");
			}
		}
		if (!sBuilder.toString().endsWith(exceptionCause)) {
			throw new Exception(sBuilder.toString().substring(0, sBuilder.toString().lastIndexOf(", ")-1));
		}
	}
	
	private void callEWSService() throws Exception {
		try {
			StringBuilder sBuilder = new StringBuilder("http://");
			sBuilder.append(cmdLineParameters.get("host"));
			sBuilder.append(":");
			sBuilder.append(cmdLineParameters.get("port"));
			sBuilder.append("/ews/services/");
			String	url1 = sBuilder.toString();
//	  	logger.debug("In CalEWS set log url " + url1);
			System.out.println("In CalEWS set url " + url1);
			System.out.flush();
			EWSClientConversation client = new EWSClientConversation();  
//	    logger.debug("In CalEWS Client Created");
			client.start(new URL(url1));  
	   
			client.authenticate(cmdLineParameters.get("user"),cmdLineParameters.get("password"));      
	 
	    
			OperationContext context = new OperationContext();  
	  
			context.setDistrict(cmdLineParameters.get("district"));
			context.setPosition(cmdLineParameters.get("position"));
			context.setMaxInstances(5);
			context.setReturnWarnings(true);

			Table tableService = client.createService(Table.class);
			
			TableServiceRetrieveRequestDTO tableRetrieveRequest = new TableServiceRetrieveRequestDTO();
			tableRetrieveRequest.setTableType("ER");
			TableServiceRetrieveReplyCollectionDTO tableRetrieveReplies =  tableService.retrieve(context, tableRetrieveRequest, null, null);
			ArrayOfTableServiceRetrieveReplyDTO retrieveReplyElements = tableRetrieveReplies.getReplyElements();
			for (TableServiceRetrieveReplyDTO reply : retrieveReplyElements.getTableServiceRetrieveReplyDTO()) {
				System.out.println("code = " + reply.getTableCode() + " description = " + reply.getCodeDescription());
			}
			client.stop(context);
			
	    }catch(EnterpriseServiceOperationException ECOS){
	  
	    	List<ErrorMessageDTO> errorList = ECOS.getFaultInfo().getErrorMessages().getErrorMessageDTO();
	    	Iterator<ErrorMessageDTO> iterator = errorList.iterator();
	    	while (iterator.hasNext()) {
	    		ErrorMessageDTO EMD = iterator.next();
	    		
	    	}
	    }
	}

}
