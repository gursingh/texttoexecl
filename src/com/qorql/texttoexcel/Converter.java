package com.qorql.texttoexcel;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.qorql.texttoexcel.exception.TextReaderException;
import com.qorql.texttoexcel.response.Response;

public class Converter {
	private static final String BASE_PATH = "D:\\Phist_Reviewed\\Processed\\path_lab_error\\";

	static Response readTextFile(String errorFileName) {
		Response response = new Response();
		try {
			List<Patient> patients = new ArrayList<Patient>();
			List<Integer> rowNumbers = new ArrayList<Integer>();
			File file = new File(BASE_PATH + errorFileName);
			BufferedReader br = new BufferedReader(new FileReader(file));
			String line = null;
			while ((line = br.readLine()) != null) {
				if (line.indexOf("Row no") != -1) {
					//rowNumbers.add(getPatientDetailRowFromLine(line));
				} else if (line.indexOf("Ignored") != -1) {
					Patient patient = getPatientDetailFromLine(line);
					patients.add(patient);
				} else {
					int rowNo = getPatientDetailRowFromLine(line);
					if (rowNo == -1) {
						Patient patient = getPatientDetailFromLine(line);
						patients.add(patient);

					} else {
						//rowNumbers.add(rowNo);
					}
				}
			}
			br.close();
			response.setPatients(patients);
			response.setRowNumbers(rowNumbers);
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
		return response;
	}

	

	static void copyRowToExcel(XSSFWorkbook wb, XSSFSheet outSheet, Row row) {
		int rowNo = outSheet.getLastRowNum() + 1;
		Row outRow = outSheet.createRow(rowNo);
		for (int i = 0; i < row.getLastCellNum(); i++) {
			Cell oldCell = row.getCell(i);
			Cell newCell = outRow.createCell(i);
			newCell.setCellType(oldCell.getCellType());

			switch (oldCell.getCellType()) {
			case Cell.CELL_TYPE_BLANK:
				newCell.setCellValue(oldCell.getStringCellValue());
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				newCell.setCellValue(oldCell.getBooleanCellValue());
				break;
			case Cell.CELL_TYPE_ERROR:
				newCell.setCellErrorValue(oldCell.getErrorCellValue());
				break;
			case Cell.CELL_TYPE_FORMULA:
				newCell.setCellFormula(oldCell.getCellFormula());
				break;
			case Cell.CELL_TYPE_NUMERIC:
				if (i == 4) {
					CellStyle cellStyle = wb.createCellStyle();
					CreationHelper createHelper = wb.getCreationHelper();
					cellStyle.setDataFormat(createHelper.createDataFormat()
							.getFormat("dd/mm/yyyy"));
					newCell.setCellStyle(cellStyle);
					newCell.setCellValue(oldCell.getNumericCellValue());
				} else
					newCell.setCellValue(oldCell.getNumericCellValue());
				break;
			case Cell.CELL_TYPE_STRING:
				newCell.setCellValue(oldCell.getRichStringCellValue());
				break;
			}

		}
	}
	public static void main(String arg[]) {
		String errorFileName = "5621ff8f441a04879c712600_error.txt";
		Response response = readTextFile(errorFileName);
		String excelFileName = "5621ff8f441a04879c712600_LalPath Patient List 3 Rev_2.xlsx";
		/*
		 * System.out.println(response.getPatients().size() +
		 * response.getRowNumbers().size());
		 */
		processExcelFile(excelFileName, response);

	}
	static void processExcelFile(String fileName, Response response) {
		try {
			FileInputStream fis = new FileInputStream(BASE_PATH + fileName);
			FileOutputStream out = new FileOutputStream(new File(BASE_PATH
					+ "5621ff8f441a04879c712600.xlsx"));
			Workbook workbook = null;
			XSSFWorkbook outWorkbook = new XSSFWorkbook();
			XSSFSheet outSheet = outWorkbook.createSheet("Sheet");
			if (fileName.toLowerCase().endsWith("xlsx")) {
				workbook = new XSSFWorkbook(fis);
			} else if (fileName.toLowerCase().endsWith("xls")) {
				workbook = new HSSFWorkbook(fis);
			}
			int numberOfSheets = workbook.getNumberOfSheets();
			int rowCount = 0;
			for (int i = 0; i < numberOfSheets; i++) {
				Sheet sheet = workbook.getSheetAt(i);
				for (int j : response.getRowNumbers()) {
					Row row = sheet.getRow(j + 1);
					copyRowToExcel(outWorkbook, outSheet, row);
					rowCount++;
				}
			}
			List<Patient> patients = response.getPatients();
			for (Patient patient : patients) {
				Row row = searchPatientInfoInExcel(workbook, patient);
				if (row != null) {
					copyRowToExcel(outWorkbook, outSheet, row);
					rowCount++;
				} else {
					System.out.println("NULL");
				}
			}
			outWorkbook.write(out);
			out.close();
			fis.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static Row searchPatientInfoInExcel(Workbook workbook,
			Patient patient) {
		int numberOfSheets = workbook.getNumberOfSheets();
		for (int i = 0; i < numberOfSheets; i++) {
			Sheet sheet = workbook.getSheetAt(i);
			for (int j = 1; j <= sheet.getLastRowNum(); j++) {
				Row row = sheet.getRow(j);
				if (row != null) {
					String name = "";
					if (row.getCell(1) != null) {
						name = row.getCell(1).getStringCellValue();
					}
					String mob = "";
					if (row.getCell(3) != null) {
						row.getCell(3).setCellType(Cell.CELL_TYPE_STRING);
						mob = row.getCell(3).getStringCellValue();
					}

					if (patient.getName().trim().equals(name.trim())) {
						if (patient.getMobileNumber() == null) {
							patient.setMobileNumber("");
						}
						if (patient.getMobileNumber().equals(mob)) {
							return row;
						}
					}
				}
			}
		}
		System.out.println(patient.getName() + "" + " G "
				+ patient.getMobileNumber());
		return null;
	}

	public static Patient getPatientDetailFromLine(String line)
			throws TextReaderException {
		try {
			String nameString = line.substring(line.indexOf("Name"));
			String mobileString = line.substring(line.indexOf("Mobile"));
			String words[] = mobileString.split(" ");
			if (words.length > 2) {
				String name = nameString.substring(5, line.indexOf("Mobile"));
				String mobileNo = words[2];
				Patient patient = new Patient();
				patient.setName(name);
				if (mobileNo.startsWith("+91")) {
					mobileNo = mobileNo.substring(3, mobileNo.length());
				}
				patient.setMobileNumber(mobileNo);
				return patient;
			} else {
				throw new TextReaderException(
						"can not get patient detail from row");
			}
		} catch (Exception e) {
			throw new TextReaderException(e.getMessage());
		}

	}

	public static int getPatientDetailRowFromLine(String line)
			throws TextReaderException {
		try {
			int rowNo = -1;
			if (line.indexOf("Row no") != -1) {
				String words[] = line.substring(line.indexOf("Row no") + 9)
						.split(" ");
				if (words.length > 0) {
					rowNo = Integer.parseInt(words[0]);
				}
			}
			return rowNo;
		} catch (Exception e) {
			throw new TextReaderException(e.getMessage());
		}
	}
}
