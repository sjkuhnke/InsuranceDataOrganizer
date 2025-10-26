package main;

public class InsuranceEntry {
	String employeeName;
	String transactionDate;
	String insuranceType;
	double amount;
	
	public InsuranceEntry(String name, String date, String type, double amt) {
		this.employeeName = name;
		this.transactionDate = date;
		this.insuranceType = type;
		this.amount = amt;
	}
}
