package org.openmf.mifos.dataimport.populator;

import java.io.IOException;

import org.openmf.mifos.dataimport.http.MifosRestClient;
import org.openmf.mifos.dataimport.populator.client.ClientWorkbookPopulator;
import org.openmf.mifos.dataimport.populator.client.OfficeSheetPopulator;
import org.openmf.mifos.dataimport.populator.loan.LoanProductSheetPopulator;
import org.openmf.mifos.dataimport.populator.loan.LoanRepaymentWorkbookPopulator;
import org.openmf.mifos.dataimport.populator.loan.LoanWorkbookPopulator;
import org.openmf.mifos.dataimport.populator.savings.SavingsProductSheetPopulator;
import org.openmf.mifos.dataimport.populator.savings.SavingsTransactionWorkbookPopulator;
import org.openmf.mifos.dataimport.populator.savings.SavingsWorkbookPopulator;

public class WorkbookPopulatorFactory {
	
	
	  public static final WorkbookPopulator createWorkbookPopulator(String parameter, String template) throws IOException {
            MifosRestClient restClient = new MifosRestClient();  
		  
	        if(template.trim().equals("client")) 
	             return new ClientWorkbookPopulator (parameter, new OfficeSheetPopulator(restClient), new PersonnelSheetPopulator(Boolean.FALSE, restClient));
	        else if(template.trim().equals("loan"))
	        	 return new LoanWorkbookPopulator(new ClientSheetPopulator(restClient), new PersonnelSheetPopulator(Boolean.TRUE, restClient),
	        			 new LoanProductSheetPopulator(restClient), new ExtrasSheetPopulator(restClient));
	        else if(template.trim().equals("loanRepaymentHistory"))
	        	 return new LoanRepaymentWorkbookPopulator(restClient, new ClientSheetPopulator(restClient), new ExtrasSheetPopulator(restClient));
	        else if(template.trim().equals("savings"))
	        	 return new SavingsWorkbookPopulator(new ClientSheetPopulator(restClient), new PersonnelSheetPopulator(Boolean.TRUE, restClient),
	        			 new SavingsProductSheetPopulator(restClient));
	        else if(template.trim().equals("savingsTransactionHistory"))
	        	 return new SavingsTransactionWorkbookPopulator(restClient, new ClientSheetPopulator(restClient), new ExtrasSheetPopulator(restClient));
	        throw new IllegalArgumentException("Can't find populator.");
	    }
}
