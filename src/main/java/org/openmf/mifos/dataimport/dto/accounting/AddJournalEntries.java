package org.openmf.mifos.dataimport.dto.accounting;

import java.util.List;
import java.util.Locale;

import org.openmf.mifos.dataimport.dto.CreditDebit;

public class AddJournalEntries {

	private final transient Integer rowIndex;

	private final String dateFormat;

	private final Locale locale;
	
	private final String officeId;
	
	private final String transactionDate;
	
	private final String currencyCode;
	
	//private final String transactionTypeId;
	
	//private final String glAccountId;
	
	//private final String amount;
	
	private final String paymentTypeId;
	
	private List<CreditDebit> debits;
	
	private List<CreditDebit> credits;

	
	public AddJournalEntries(String officeId, String transactionDate,
			String currencyCode, String paymentTypeId,Integer rowIndex, List<CreditDebit> credits,
			List<CreditDebit> debits) {
		
		this.officeId = officeId;
		this.transactionDate = transactionDate;
		this.rowIndex = rowIndex;
		this.currencyCode = currencyCode;
		//this.transactionTypeId = transactionTypeId;
		//this.glAccountId= glAccountId;
		//this.amount = amount;
		this.paymentTypeId= paymentTypeId;
		this.credits = credits;
		this.debits = debits;
		this.dateFormat = "dd MMMM yyyy";
		this.locale = Locale.ENGLISH;

	}

	public Locale getLocale() {
		return locale;
	}

	public Integer getRowIndex() {
		return rowIndex;
	}

	public String getDateFormat() {
		return dateFormat;
	}

	public String getOfficeId() {
		return officeId;
	}

	public String getTransactionDate() {
		return transactionDate;
	}

	public String getCurrencyCode() {
		return currencyCode;
	}

	public List<CreditDebit> getCredits() {
		return credits;
	}

	public List<CreditDebit> getDebits() {
		return debits;
	}

	//public String getGlAccountId() {
	//	return glAccountId;
	//}

	// String getAmount() {
	//	return amount;
	//}

	//public String getTransactionTypeId() {
	//	return transactionTypeId;
	//}

	public String getPaymentTypeId() {
		return paymentTypeId;
	}

	
	public void addDebits(CreditDebit debit) {
		this.debits.add(debit);
	}

	public void addCredits(CreditDebit credit) {
		this.credits.add(credit);
	}


}
