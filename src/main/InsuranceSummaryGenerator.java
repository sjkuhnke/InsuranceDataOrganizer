package main;

import java.io.*;
import java.util.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class InsuranceSummaryGenerator {
	
	public static void main(String[] args) {
		String inputFile = "Transaction_Report.xlsx";
		String outputFile = "Insurance_Summary.xlsx";
		
		// TODO:
		// Horizontal Sums
		// Vertical Sums
		// Display the =val1+val2 for Health Insurance instead of just the sum
		
		try {
			processInsuranceData(inputFile, outputFile);
			System.out.println("Insurance summary generated successfully!");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void processInsuranceData(String inputFile, String outputFile) throws IOException {
		// input
		FileInputStream fis = new FileInputStream(inputFile);
		Workbook inputWorkbook = WorkbookFactory.create(fis);
		Sheet inputSheet = inputWorkbook.getSheetAt(0);
		
		// data structures!
		Map<String, List<InsuranceEntry>> insuranceByType = new LinkedHashMap<>();
		Set<String> allDates = new LinkedHashSet<>();
		Set<String> allEmployees = new LinkedHashSet<>();
		
		System.out.println("Starting to read data...");
		
		// read starting from row 3 (could change, this is just to be safe)
		for (int i = 3; i <= inputSheet.getLastRowNum(); i++) {
			Row row = inputSheet.getRow(i);
			if (row == null) continue;
			
			// column C for Transaction Type
			Cell transactionTypeCell = row.getCell(2);
			if (transactionTypeCell == null || !getCellValue(transactionTypeCell).equals("Payroll Check")) {
				continue;
			}
			
			// data
			String date = getCellValue(row.getCell(1)); // column B
			String employeeName = getCellValue(row.getCell(4)); // column E
			String memoDesc = getCellValue(row.getCell(5)); // column F
			double amount = getNumericValue(row.getCell(8)); // column I
			
			if (amount == 0.0 || memoDesc.isEmpty()) continue;
			
			// categorize insurance types
			String insuranceType = categorizeInsurance(memoDesc);
			if (insuranceType == null) continue;
			
			// store entry
			InsuranceEntry entry = new InsuranceEntry(employeeName, date, insuranceType, amount);
			insuranceByType.computeIfAbsent(insuranceType, k -> new ArrayList<>()).add(entry);
			allDates.add(date);
			allEmployees.add(employeeName);
		}
		
		System.out.println("Found " + insuranceByType.size() + " insurance types");
		System.out.println("Insurance types: " + insuranceByType.keySet());
		System.out.println("Total dates: " + allDates.size());
		System.out.println("Total employees: " + allEmployees.size());
		
		fis.close();
		inputWorkbook.close();
		
		// failsafe: don't create empty workbooks
		if (insuranceByType.isEmpty()) {
			System.err.println("WARNING: No insurance data found! Workbook not created.");
			return;
		}
		
		// create output workbook
		Workbook outputWorkbook = new XSSFWorkbook();
		
		// sheet for each insurance type
		for (String insuranceType : insuranceByType.keySet()) {
			System.out.println("\nCreating sheet for: " + insuranceType);
			System.out.println("  Entries: " + insuranceByType.get(insuranceType).size());
			
			try {
				createInsuranceSheet(outputWorkbook, insuranceType, insuranceByType.get(insuranceType), allDates, allEmployees);
				System.out.println("  Sheet created successfully");
			} catch (Exception e) {
				System.out.println("  ERROR creating sheet: " + e.getMessage());
				e.printStackTrace();
			}
		}
		
		System.out.println("\nWriting output file...");
		
		// write output file
		FileOutputStream fos = new FileOutputStream(outputFile);
		outputWorkbook.write(fos);
		fos.close();
		outputWorkbook.close();
		
		System.out.println("Output file written");
	}
	
	private static String categorizeInsurance(String memoDesc) {
		if (memoDesc.contains("Health Insurance") && !memoDesc.contains("S-Corp")) {
			return "Health Insurance";
		} else if (memoDesc.contains("Dental Insurance")) {
			return "Dental Insurance";
		} else if (memoDesc.contains("Insurance - Advantage Group")) {
			return "Advantage Group Insurance";
		} else if (memoDesc.contains("Federal Unemployment")) {
			return "Federal Unemployment";
		} else if (memoDesc.contains("WI SUI Employer")) {
			return "WI SUI Employer";
		}
		return null;
	}
	
	private static void createInsuranceSheet(Workbook wb, String insuranceType, 
		List<InsuranceEntry> entries,
		Set<String> allDates, 
		Set<String> allEmployees) {
		
		Sheet sheet = wb.createSheet(sanitizeSheetName(insuranceType));
		
		Map<String, Map<String, List<Double>>> employeeData = new LinkedHashMap<>();
		List<String> dateList = new ArrayList<>(allDates);
		List<String> employeeList = new ArrayList<>(allEmployees);
		
		Set<String> relevantDates = new LinkedHashSet<>();
		Set<String> relevantEmployees = new LinkedHashSet<>();
		
		// init structure for employees
		for (String emp : employeeList) {
			employeeData.put(emp, new LinkedHashMap<>());
			for (String date : dateList) {
				employeeData.get(emp).put(date, new ArrayList<>());
			}
		}
		
		// populate with actual data
		for (InsuranceEntry entry : entries) {
			if (entry.insuranceType.equals(insuranceType)) {
				employeeData.get(entry.employeeName).get(entry.transactionDate).add(entry.amount);
				relevantDates.add(entry.transactionDate);
				relevantEmployees.add(entry.employeeName);
			}
		}
		
		// relevant dates only
		List<String> relevantDateList = new ArrayList<>(relevantDates);
		
		Row headerRow = sheet.createRow(0);
		headerRow.createCell(0).setCellValue(insuranceType);
		
		int currentRow = 1;
		
		// create rows per employee
		for (String employee : employeeList) {
			Row row = sheet.createRow(currentRow);
			int currentCol = 0;
			
			// for each date create Amount, Name, Date columns
			for (String date : relevantDateList) {
				List<Double> amounts = employeeData.get(employee).get(date);
				
				if (!amounts.isEmpty()) {
					// amt cell - sum all amounts as a single value
					Cell amountCell = row.createCell(currentCol);
					double total = 0.0;
					for (double amt : amounts) {
						total += amt;
					}
					amountCell.setCellValue(total);
					
					// name
					row.createCell(currentCol + 1).setCellValue(employee);
					
					// date
					row.createCell(currentCol + 2).setCellValue(date);
				}
				// move to next date group (4 cols)
				currentCol += 4;
			}
			currentRow++;
		}
		
		// total row
		Row totalRow = sheet.createRow(currentRow);
		int currentCol = 0;
		
		for (String date : relevantDateList) {
			totalRow.createCell(currentCol).setCellValue("Total");
			// move to next date group
			currentCol += 4;
		}
	}
	
	private static String sanitizeSheetName(String name) {
		// sheet names can't exceed 31 chars
		return name.replaceAll("[\\\\/*?\\[\\]:]", "").substring(0, Math.min(31, name.length()));
	}
	
	private static String getCellValue(Cell cell) {
		if (cell == null) return "";
		switch (cell.getCellType()) {
		case STRING:
			return cell.getStringCellValue().trim();
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				return cell.getDateCellValue().toString();
			}
			return String.valueOf((int)cell.getNumericCellValue());
		default:
			return "";
		}
	}
	
	private static double getNumericValue(Cell cell) {
		if (cell == null) return 0.0;
		try {
			return cell.getNumericCellValue();
		} catch (Exception e) {
			return 0.0;
		}
	}
}
