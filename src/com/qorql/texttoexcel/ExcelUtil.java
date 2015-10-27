package com.qorql.texttoexcel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil {
	public static void copyRowToExcel(XSSFWorkbook wb, XSSFSheet outSheet,
			Row row, boolean isLabFile) {
		int rowNo = outSheet.getLastRowNum() + 1;
		Row outRow = outSheet.createRow(rowNo);
		for (int i = 0; i < row.getLastCellNum(); i++) {
			Cell oldCell = row.getCell(i);
			Cell newCell = outRow.createCell(i);
			if (oldCell != null) {
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
					if (isLabFile) {
						if (i == 4) {
							CellStyle cellStyle = wb.createCellStyle();
							CreationHelper createHelper = wb
									.getCreationHelper();
							cellStyle
									.setDataFormat(createHelper
											.createDataFormat().getFormat(
													"dd/mm/yyyy"));
							newCell.setCellStyle(cellStyle);
							newCell.setCellValue(oldCell.getNumericCellValue());
						} else
							newCell.setCellValue(oldCell.getNumericCellValue());
					} else {
						if (i == 1 || i == 5) {
							CellStyle cellStyle = wb.createCellStyle();
							CreationHelper createHelper = wb
									.getCreationHelper();
							cellStyle
									.setDataFormat(createHelper
											.createDataFormat().getFormat(
													"dd/mm/yyyy"));
							newCell.setCellStyle(cellStyle);
							newCell.setCellValue(oldCell.getNumericCellValue());
						} else
							newCell.setCellValue(oldCell.getNumericCellValue());
					}

					break;
				case Cell.CELL_TYPE_STRING:
					newCell.setCellValue(oldCell.getRichStringCellValue());
					break;
				}

			}

		}
	}
}
