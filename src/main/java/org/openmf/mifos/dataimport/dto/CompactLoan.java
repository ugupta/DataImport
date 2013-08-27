package org.openmf.mifos.dataimport.dto;

import java.util.Comparator;

public class CompactLoan {

	private final String accountNo;
	
	private final String clientName;
	
	private final String loanProductName;
	
	private final Double principal;
	
	private final Status status;
	
	public CompactLoan(String accountNo, String clientName, String loanProductName, Double principal, Status status) {
		this.accountNo = accountNo;
		this.clientName = clientName;
		this.loanProductName = loanProductName;
		this.principal = principal;
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
	
	public static Comparator<CompactLoan> ClientNameComparator = new Comparator<CompactLoan>() {
		
	@Override
	public int compare(CompactLoan loan1, CompactLoan loan2) {
		String clientOfLoan1 = loan1.getClientName().toUpperCase();
		String clientOfLoan2 = loan2.getClientName().toUpperCase(); 
		return clientOfLoan1.compareTo(clientOfLoan2);
	 }
	};
}
