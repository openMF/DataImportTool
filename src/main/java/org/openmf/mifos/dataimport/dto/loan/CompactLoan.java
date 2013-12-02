package org.openmf.mifos.dataimport.dto.loan;

import java.util.Comparator;
import java.util.Locale;

import org.openmf.mifos.dataimport.dto.Status;

public class CompactLoan {

	private final String accountNo;
	
	private final String clientName;
	
	private final String loanProductName;
	
	private final Double principal;
	
	private final LoanTimeline timeline;
	
	private final Status status;
	
	public CompactLoan(String accountNo, String clientName, String loanProductName, Double principal, LoanTimeline timeline, Status status) {
		this.accountNo = accountNo;
		this.clientName = clientName;
		this.loanProductName = loanProductName;
		this.principal = principal;
		this.timeline = timeline;
		this.status = status;
	}
	
	public String getAccountNo() {
        return this.accountNo;
    }
	
	public String getClientName() {
        return this.clientName;
    }
	
	public String getLoanProductName() {
        return this.loanProductName;
    }
	
	public Double getPrincipal() {
		return this.principal;
	}
	
	public Boolean isActive() {
		return this.status.isActive();
	}
	
	public LoanTimeline getTimeline() {
		return timeline;
	}

	public static final Comparator<CompactLoan> ClientNameComparator = new Comparator<CompactLoan>() {
		
	@Override
	public int compare(CompactLoan loan1, CompactLoan loan2) {
		String clientOfLoan1 = loan1.getClientName().toUpperCase(Locale.ENGLISH);
		String clientOfLoan2 = loan2.getClientName().toUpperCase(Locale.ENGLISH); 
		return clientOfLoan1.compareTo(clientOfLoan2);
	 }
	};
}
