package org.openmf.mifos.dataimport.populator;

import java.io.IOException;

import org.openmf.mifos.dataimport.http.MifosRestClient;
import org.openmf.mifos.dataimport.populator.accounting.AddJournalEntriesWorkbookPopulator;
import org.openmf.mifos.dataimport.populator.client.CenterWorkbookPopulator;
import org.openmf.mifos.dataimport.populator.client.ClientWorkbookPopulator;
import org.openmf.mifos.dataimport.populator.client.GroupWorkbookPopulator;
import org.openmf.mifos.dataimport.populator.loan.AddGuarantorWorkbookPopulator;
import org.openmf.mifos.dataimport.populator.loan.LoanProductSheetPopulator;
import org.openmf.mifos.dataimport.populator.loan.LoanRepaymentWorkbookPopulator;
import org.openmf.mifos.dataimport.populator.loan.LoanWorkbookPopulator;
import org.openmf.mifos.dataimport.populator.savings.ClosingOfSavingsAccountsWorkbookPopulator;
import org.openmf.mifos.dataimport.populator.savings.FixedDepositProductSheetPopulator;
import org.openmf.mifos.dataimport.populator.savings.FixedDepositWorkbookPopulator;
import org.openmf.mifos.dataimport.populator.savings.RecurringDepositProductSheetPopulator;
import org.openmf.mifos.dataimport.populator.savings.RecurringDepositTransactionWorkbookPopulator;
import org.openmf.mifos.dataimport.populator.savings.RecurringDepositWorkbookPopulator;
import org.openmf.mifos.dataimport.populator.savings.SavingsProductSheetPopulator;
import org.openmf.mifos.dataimport.populator.savings.SavingsTransactionWorkbookPopulator;
import org.openmf.mifos.dataimport.populator.savings.SavingsWorkbookPopulator;

public class WorkbookPopulatorFactory {
	
	
	  public static final WorkbookPopulator createWorkbookPopulator(String parameter, String template) throws IOException {
            MifosRestClient restClient = new MifosRestClient();  
		  
	        if(template.trim().equals("client")) 
	             return new ClientWorkbookPopulator (parameter, new OfficeSheetPopulator(restClient), new PersonnelSheetPopulator(Boolean.FALSE, restClient));
	        else if(template.trim().equals("groups"))
	        	 return new GroupWorkbookPopulator(new OfficeSheetPopulator(restClient), new PersonnelSheetPopulator(Boolean.FALSE, restClient), new CenterSheetPopulator(restClient),
	        			 new ClientSheetPopulator(restClient));
	        else if(template.trim().equals("centers"))
	        	 return new CenterWorkbookPopulator(new OfficeSheetPopulator(restClient), new PersonnelSheetPopulator(Boolean.FALSE, restClient));
	        else if(template.trim().equals("loan"))
	        	 return new LoanWorkbookPopulator(new OfficeSheetPopulator(restClient), new ClientSheetPopulator(restClient), new GroupSheetPopulator(restClient),
	        			 new PersonnelSheetPopulator(Boolean.TRUE, restClient), new LoanProductSheetPopulator(restClient), new ExtrasSheetPopulator(restClient));
	        else if(template.trim().equals("loanRepaymentHistory"))
	        	 return new LoanRepaymentWorkbookPopulator(restClient, new OfficeSheetPopulator(restClient), new ClientSheetPopulator(restClient), new ExtrasSheetPopulator(restClient));
	        else if(template.trim().equals("savings"))
	        	 return new SavingsWorkbookPopulator(new OfficeSheetPopulator(restClient), new ClientSheetPopulator(restClient), new GroupSheetPopulator(restClient),
	        			 new PersonnelSheetPopulator(Boolean.TRUE, restClient), new SavingsProductSheetPopulator(restClient));
	        else if(template.trim().equals("savingsTransactionHistory"))
	        	 return new SavingsTransactionWorkbookPopulator(restClient, new OfficeSheetPopulator(restClient), new ClientSheetPopulator(restClient), new ExtrasSheetPopulator(restClient));
	        else if(template.trim().equals("fixedDeposit"))
	        	 return new FixedDepositWorkbookPopulator(new OfficeSheetPopulator(restClient), new ClientSheetPopulator(restClient),
	        			 new PersonnelSheetPopulator(Boolean.TRUE, restClient), new FixedDepositProductSheetPopulator(restClient));
	        else if(template.trim().equals("recurringDeposit"))
	        	return new RecurringDepositWorkbookPopulator(new OfficeSheetPopulator(restClient), new ClientSheetPopulator(restClient),
	        			 new PersonnelSheetPopulator(Boolean.TRUE, restClient), new RecurringDepositProductSheetPopulator(restClient));
	        else if(template.trim().equals("recurringDepositHistory"))
	        	 return new RecurringDepositTransactionWorkbookPopulator(restClient, new OfficeSheetPopulator(restClient), new ClientSheetPopulator(restClient), new ExtrasSheetPopulator(restClient));
	        else if(template.trim().equals("closingOfSavingsAccounts"))
	        	 return new ClosingOfSavingsAccountsWorkbookPopulator(restClient, new OfficeSheetPopulator(restClient), new ClientSheetPopulator(restClient));
	        else if(template.trim().equals("journalentries"))
	        	return new AddJournalEntriesWorkbookPopulator(new OfficeSheetPopulator(restClient), new GlAccountSheetPopulator(restClient),new ExtrasSheetPopulator(restClient));
	        else if(template.trim().equals("guarantor"))
	        	
	        return new AddGuarantorWorkbookPopulator(restClient, new OfficeSheetPopulator(restClient), new ClientSheetPopulator(restClient));
	        throw new IllegalArgumentException("Can't find populator.");
	    }
}
