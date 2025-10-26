package main;

import java.awt.FileDialog;
import java.awt.Frame;

import java.io.*;
import java.util.*;

import javax.swing.JOptionPane;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class InsuranceSummaryGenerator {
	
	public static void main(String[] args) {
		
		// INPUT
		FileDialog fileChooser = new FileDialog((Frame) null, "Select Transaction Report Excel File", FileDialog.LOAD);
		fileChooser.setFile("*.xlsx;*.xls");
		fileChooser.setVisible(true);
		
		String inputFile = fileChooser.getFile();
		String inputDirectory = fileChooser.getDirectory();
		
		if (inputFile != null && inputDirectory != null) {
			String fullInputPath = inputDirectory + inputFile;
			System.out.println("Selected input file: " + fullInputPath);
			
			// OUTPUT
			FileDialog saveChooser = new FileDialog((Frame) null, "Save Insurance Summary As", FileDialog.SAVE);
			saveChooser.setFile("Insurance_Summary.xlsx");
			saveChooser.setVisible(true);
			
			String outputFile = saveChooser.getFile();
			String outputDirectory = saveChooser.getDirectory();

			if (outputFile != null && outputDirectory != null) {
				String fullOutputPath = outputDirectory + outputFile;
				
				if (!fullOutputPath.toLowerCase().endsWith(".xlsx")) {
					fullOutputPath += ".xlsx";
				}
				
				System.out.println("Output file will be saved to: " + fullOutputPath);
				
				try {
					processInsuranceData(fullInputPath, fullOutputPath);
					System.out.println("Insurance summary generated successfully!");
				} catch (Exception e) {
					JOptionPane.showMessageDialog(null, e.getMessage());
					e.printStackTrace();
					System.exit(1);
				}
				System.exit(0);
			} else {
				System.out.println("Save location selection cancelled.");
				JOptionPane.showMessageDialog(null, "Save location selection cancelled.");
				System.exit(0);
			}
		} else {
			System.out.println("Input file selection cancelled.");
			JOptionPane.showMessageDialog(null, "Input file selection cancelled.");
			System.exit(0);
		}
	}

	public static void processInsuranceData(String inputFile, String outputFile) throws IOException {		
		// data structures!
		Map<String, List<InsuranceEntry>> insuranceByType = new LinkedHashMap<>();
		Set<String> allDates = new LinkedHashSet<>();
		Set<String> allEmployees = new LinkedHashSet<>();
		
		try (FileInputStream fis = new FileInputStream(inputFile);
			Workbook inputWorkbook = WorkbookFactory.create(fis)) {
			
			Sheet inputSheet = inputWorkbook.getSheetAt(0);
			
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
		}
		
		// failsafe: don't create empty workbooks
		if (insuranceByType.isEmpty()) {
			System.err.println("WARNING: No insurance data found! Workbook not created.");
			JOptionPane.showMessageDialog(null, "WARNING: No insurance data found! Workbook not created.");
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
				JOptionPane.showMessageDialog(null, e.getMessage());
				e.printStackTrace();
			}
		}
		
		System.out.println("\nWriting output file...");
		
		// write output file
		try (FileOutputStream fos = new FileOutputStream(outputFile)) {
			outputWorkbook.write(fos);
		} finally {
			outputWorkbook.close();
		}
		
		System.out.println("Output file written");
		JOptionPane.showMessageDialog(null, "Data succesfully written to " + outputFile + "!");
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
		
		CellStyle accountingStyle = wb.createCellStyle();
		DataFormat format = wb.createDataFormat();
		accountingStyle.setDataFormat(format.getFormat("_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)"));
		
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
		int dataStartCol = 5;
		
		// create rows per employee
		for (String employee : employeeList) {
			Row row = sheet.createRow(currentRow);
			int currentCol = dataStartCol;
			
			row.createCell(2).setCellValue(employee);
			
			// for each date create Amount, Name, Date columns
			for (String date : relevantDateList) {
				List<Double> amounts = employeeData.get(employee).get(date);
				
				if (!amounts.isEmpty()) {
					// amt cell
					Cell amountCell = row.createCell(currentCol);
					amountCell.setCellStyle(accountingStyle);
					if (amounts.size() == 1) {
						amountCell.setCellValue(amounts.get(0));
					} else {
						StringBuilder formula = new StringBuilder();
						for (int i = 0; i < amounts.size(); i++) {
							if (i > 0) formula.append("+");
							formula.append(amounts.get(i));
						}
						amountCell.setCellFormula(formula.toString());
					}
					
					// name
					row.createCell(currentCol + 1).setCellValue(employee);
					
					// date
					row.createCell(currentCol + 2).setCellValue(date);
				}
				// move to next date group (4 cols)
				currentCol += 4;
			}
			
			currentRow++;
			
			if (!relevantDateList.isEmpty()) {
				int startCol = dataStartCol;
				int endCol = dataStartCol + ((relevantDateList.size() - 1) * 4);
				String formula = String.format("SUM(%s%d:%s%d)", getColumnLetter(startCol), currentRow, getColumnLetter(endCol), currentRow);
				Cell totalCell = row.createCell(3);
				totalCell.setCellStyle(accountingStyle);
				totalCell.setCellFormula(formula);
			}
		}
		
		// total row
		Row totalRow = sheet.createRow(currentRow);
		int currentCol = dataStartCol;
		
		for (int i = 0; i < relevantDateList.size(); i++) {
			String columnLetter = getColumnLetter(currentCol);
			String formula = String.format("SUM(%s2:%s%d)", columnLetter, columnLetter, currentRow);
			Cell totalCell = totalRow.createCell(currentCol);
			totalCell.setCellStyle(accountingStyle);
			totalCell.setCellFormula(formula);
			
			// move to next date group
			currentCol += 4;
		}
		
		// grand total in horizontal sum section
		totalRow.createCell(2).setCellValue("Total");
		String totalFormula = String.format("SUM(D2:D%d)", currentRow);
		Cell grandTotalCell = totalRow.createCell(3);
		grandTotalCell.setCellStyle(accountingStyle);
		grandTotalCell.setCellFormula(totalFormula);
		
		sheet.autoSizeColumn(2);
		sheet.setColumnWidth(3, 110 * 37);
		
		currentCol = dataStartCol;
		for (int i = 0; i < relevantDateList.size(); i++) {
			sheet.setColumnWidth(currentCol, 92 * 37);
			sheet.autoSizeColumn(currentCol + 1);
			sheet.autoSizeColumn(currentCol + 2);
			
			currentCol += 4;
	    }
	}
	
	private static String getColumnLetter(int col) {
		StringBuilder result = new StringBuilder();
		while (col >= 0) {
			result.insert(0, (char)('A' + (col % 26)));
			col = (col / 26) - 1;
		}
		return result.toString();
	}
	
	@SuppressWarnings("unused")
	private static String getCellReference(int row, int col) {
		return getColumnLetter(col) + (row + 1);
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
