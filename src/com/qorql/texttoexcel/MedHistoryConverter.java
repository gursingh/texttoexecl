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
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.qorql.texttoexcel.exception.TextReaderException;
import com.qorql.texttoexcel.response.Response;
import com.qorql.texttoexcel.response.RowNumberResponse;

public class MedHistoryConverter {
	private static final String BASE_PATH = Consts.MEDI_HISTORY_FILE_BASE;

	static Response readTextFile(String errorFileName) {
		Response response = new Response();
		try {
			List<Patient> patients = new ArrayList<Patient>();
			List<RowNumberResponse> rowNumbers = new ArrayList<RowNumberResponse>();
			File file = new File(BASE_PATH + errorFileName);
			BufferedReader br = new BufferedReader(new FileReader(file));
			String line = null;
			while ((line = br.readLine()) != null) {
				if (line.indexOf("Row Number") != -1) {
					RowNumberResponse rowNumberResponse = getPatientDetailRowFromLine(line);
					if (rowNumberResponse != null)
						rowNumbers.add(rowNumberResponse);
				} else if (line.indexOf("Ignored") != -1) {
					Patient patient = getPatientDetailFromLine(line);
					patients.add(patient);
				} else {
					RowNumberResponse rowNumberResponse = getPatientDetailRowFromLine(line);
					if (rowNumberResponse != null)
						rowNumbers.add(rowNumberResponse);
					else {
						Patient patient = getPatientDetailFromLine(line);
						patients.add(patient);
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

	public static void main(String arg[]) {
		String errorFileName = Consts.MEDI_HISTORY_ERROR_FILE;
		Response response = readTextFile(errorFileName);
		String excelFileName = Consts.MEDI_HISTORY_INPUT_FILE;
		processExcelFile(excelFileName, response);

	}

	static void processExcelFile(String fileName, Response response) {
		try {
			FileInputStream fis = new FileInputStream(BASE_PATH + fileName);
			FileOutputStream out = new FileOutputStream(new File(BASE_PATH
					+ Consts.MEDI_HISTORY_OUTPUT_FILE));
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
				for (RowNumberResponse rowNumberResponse : response
						.getRowNumbers()) {
					Row row = sheet.getRow(rowNumberResponse.getRowNo());
					Cell cell = row.createCell(row.getLastCellNum() + 1);
					cell.setCellType(Cell.CELL_TYPE_STRING);
					cell.setCellValue(rowNumberResponse.getMessage());
					ExcelUtil.copyRowToExcel(outWorkbook, outSheet, row, false);
					rowCount++;
				}
			}
			List<Patient> patients = response.getPatients();
			for (Patient patient : patients) {
				Row row = searchPatientInfoInExcel(workbook, patient);
				if (row != null) {
					Cell cell = row.createCell(row.getLastCellNum() + 1);
					cell.setCellType(Cell.CELL_TYPE_STRING);
					cell.setCellValue(patient.getErrorMessage());
					ExcelUtil.copyRowToExcel(outWorkbook, outSheet, row, false);
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
					if (row.getCell(2) != null) {
						name = row.getCell(2).getStringCellValue();
					}
					String mob = "";
					if (row.getCell(4) != null) {
						row.getCell(4).setCellType(Cell.CELL_TYPE_STRING);
						mob = row.getCell(4).getStringCellValue();
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
		return null;
	}

	public static Patient getPatientDetailFromLine(String line)
			throws TextReaderException {
		try {
			String nameString = line.substring(line.indexOf("Name"));
			String mobileString = line.substring(line.indexOf("Mobile"));
			String words[] = mobileString.split(" ");
			if (words.length > 2) {
				Patient patient = new Patient();
				if (line.indexOf("error") != -1) {
					patient.setErrorMessage(line.substring(line
							.indexOf("error")));
				}
				String name = nameString.substring(5, line.indexOf("Mobile"));
				String mobileNo = words[2];
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

	public static RowNumberResponse getPatientDetailRowFromLine(String line)
			throws TextReaderException {
		try {
			RowNumberResponse rowNumberResponse = null;
			int rowNo = -1;
			if (line.indexOf("Row Number") != -1) {
				String words[] = line
						.substring(line.indexOf("Row Number") + 11).split(" ");
				if (words.length > 0) {
					rowNumberResponse = new RowNumberResponse();
					rowNo = Integer.parseInt(words[0]);
					rowNumberResponse.setRowNo(rowNo);
					if (line.indexOf("error") != -1) {
						rowNumberResponse.setMessage(line.substring(line
								.indexOf("error")));
					} else {
						rowNumberResponse.setMessage("");
					}

				}
			}
			return rowNumberResponse;
		} catch (Exception e) {
			throw new TextReaderException(e.getMessage());
		}
	}
}
