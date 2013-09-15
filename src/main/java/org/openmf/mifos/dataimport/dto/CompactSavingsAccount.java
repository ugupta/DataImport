package org.openmf.mifos.dataimport.dto;

import java.util.Comparator;
import java.util.Locale;

public class CompactSavingsAccount {

private final String accountNo;
	
	private final String clientName;
	
	private final String savingsProductName;
	
	private final Double minRequiredOpeningBalance;
	
	private final Status status;
	
	public CompactSavingsAccount(String accountNo, String clientName, String savingsProductName, Double minRequiredOpeningBalance, Status status) {
		this.accountNo = accountNo;
		this.clientName = clientName;
		this.savingsProductName = savingsProductName;
		this.minRequiredOpeningBalance = minRequiredOpeningBalance;
		this.status = status;
	}

	public String getClientName() {
		return clientName;
	}

	public String getAccountNo() {
		return accountNo;
	}

	public String getSavingsProductName() {
		return savingsProductName;
	}

	public Double getMinRequiredOpeningBalance() {
		return minRequiredOpeningBalance;
	}

	public Boolean isActive() {
		return this.status.isActive();
	}
	
	public static final Comparator<CompactSavingsAccount> ClientNameComparator = new Comparator<CompactSavingsAccount>() {
		
		@Override
		public int compare(CompactSavingsAccount savings1, CompactSavingsAccount savings2) {
			String clientOfSavings1 = savings1.getClientName().toUpperCase(Locale.ENGLISH);
			String clientOfSavings2 = savings2.getClientName().toUpperCase(Locale.ENGLISH); 
			return clientOfSavings1.compareTo(clientOfSavings2);
		 }
		};
}
