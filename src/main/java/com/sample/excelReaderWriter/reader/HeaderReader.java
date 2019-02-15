package com.sample.excelReaderWriter.reader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Collections;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.IndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.ResourceUtils;

import com.sample.excelReaderWriter.entity.SheetHeaderMasterTest;
import com.sample.excelReaderWriter.entity.SheetMasterTest;
import com.sample.excelReaderWriter.entity.TemplateMasterTest;

public class HeaderReader {

	public static void main(String[] args) throws IOException {

		// For maintaining unique indexes
		long templateIndex = 0;
		long sheetIndex = 0;
		long headerIndex = 0;

		// Parent entity containing template, sheet and header details
		TemplateMasterTest templateMasterTest = new TemplateMasterTest();

		// Test template to be created with headers
		File templateFile = ResourceUtils.getFile("classpath:templates/test.xlsx");
		try (FileInputStream fis = new FileInputStream(templateFile)) {

			try (XSSFWorkbook xlsxWorkbook = new XSSFWorkbook(fis)) {

				int numOfSheets = xlsxWorkbook.getNumberOfSheets();

				templateMasterTest.setTemplateId(templateIndex++);
				templateMasterTest.setTemplateName("Test");

				// Set for filling all sheet details within a template
				Set<SheetMasterTest> sheetMasterTestSet = templateMasterTest.getSheetMasterTestList();

				if (numOfSheets > 0) {
					for (int count = 0; count < numOfSheets; count++) {

						XSSFSheet currentSheet = xlsxWorkbook.getSheetAt(count);

						// Filling sheet information
						SheetMasterTest sheetMasterTest = new SheetMasterTest();
						sheetMasterTest.setSheetId(sheetIndex);
						sheetMasterTest.setSheetIndex(count);
						sheetMasterTest.setSheetName(currentSheet.getSheetName());

						// Set to contain header details for current sheet
						Set<SheetHeaderMasterTest> sheetHeaderMasterTestSet = sheetMasterTest
								.getSheetHeaderMasterTestList();

						// Extracting the merged cell information of current sheet
						List<CellRangeAddress> mergedCellList = currentSheet.getMergedRegions();

						Iterator<Row> rowIterator = currentSheet.iterator();
						while (rowIterator.hasNext()) {
							XSSFRow row = (XSSFRow) rowIterator.next();

							Iterator<Cell> cellIterator = row.cellIterator();
							while (cellIterator.hasNext()) {
								XSSFCell cell = (XSSFCell) cellIterator.next();

								// Only taking header details of unmerged cells
								if (!checkCellMerged(mergedCellList, cell.getRowIndex(), cell.getColumnIndex())) {

									SheetHeaderMasterTest sheetHeaderMasterTest = new SheetHeaderMasterTest();
									sheetHeaderMasterTest.setSheetId(sheetIndex);
									sheetHeaderMasterTest.setHeaderId(headerIndex);
									sheetHeaderMasterTest.setColFrom(cell.getColumnIndex());
									sheetHeaderMasterTest.setColTo(cell.getColumnIndex());
									sheetHeaderMasterTest.setRowFrom(cell.getRowIndex());
									sheetHeaderMasterTest.setRowTo(cell.getRowIndex());
									sheetHeaderMasterTest.setHeaderText(cell.getStringCellValue());
									sheetHeaderMasterTest.setHeaderColor(cell.getCellStyle().getFillForegroundColorColor().getARGBHex());

									// Adding each unmerged header details to header set of sheet
									sheetHeaderMasterTestSet.add(sheetHeaderMasterTest);
									headerIndex++;
								}

							}
						}

						// Fetching header details of merged cells
						Set<SheetHeaderMasterTest> mergedCellHeaders = getMergedCellHeaders(currentSheet, sheetIndex,
								headerIndex);
						// Adding all merged header details to header set of sheet
						sheetHeaderMasterTestSet.addAll(mergedCellHeaders);

						// Adding all headers to sheet header set (seems redundant. Do Verify)
						sheetMasterTest.setSheetHeaderMasterTestList(sheetHeaderMasterTestSet);

						// Adding sheet details to the sheet set
						sheetMasterTestSet.add(sheetMasterTest);
						sheetIndex++;

					}
					// Adding sheet details to the sheet set of template (seems redundant. Do
					// Verify)
					templateMasterTest.setSheetMasterTestList(sheetMasterTestSet);
				}

				xlsxWorkbook.close();
			}
			fis.close();
		}

		System.out.println(templateMasterTest);
		// Reproduce template from template entity
		writeExcelTemplate(templateMasterTest);
		System.out.println("End");

	}

	public static boolean checkCellMerged(List<CellRangeAddress> mergedCellList, int rowInd, int colInd) {
		boolean merged = false;

		for (CellRangeAddress cellRangeAddress : mergedCellList) {
			if (cellRangeAddress.containsRow(rowInd) && cellRangeAddress.containsColumn(colInd)) {
				merged = true;
				break;
			}
		}
		return merged;
	}

	public static Set<SheetHeaderMasterTest> getMergedCellHeaders(XSSFSheet currentSheet, long sheetIndex,
			long headerIndex) {
		Set<SheetHeaderMasterTest> mergedCellHeaders = new HashSet<SheetHeaderMasterTest>();
		List<CellRangeAddress> mergedCellList = currentSheet.getMergedRegions();
		for (CellRangeAddress cellRangeAddress : mergedCellList) {

			SheetHeaderMasterTest sheetHeaderMasterTest = new SheetHeaderMasterTest();
			sheetHeaderMasterTest.setSheetId(sheetIndex);
			sheetHeaderMasterTest.setHeaderId(headerIndex);
			sheetHeaderMasterTest.setColFrom(cellRangeAddress.getFirstColumn());
			sheetHeaderMasterTest.setColTo(cellRangeAddress.getLastColumn());
			sheetHeaderMasterTest.setRowFrom(cellRangeAddress.getFirstRow());
			sheetHeaderMasterTest.setRowTo(cellRangeAddress.getLastRow());
			sheetHeaderMasterTest.setHeaderColor(currentSheet.getRow(cellRangeAddress.getFirstRow())
					.getCell(cellRangeAddress.getFirstColumn()).getCellStyle().getFillForegroundColorColor().getARGBHex());
			// Data of merged cells is always present in the first cell
			sheetHeaderMasterTest.setHeaderText(currentSheet.getRow(cellRangeAddress.getFirstRow())
					.getCell(cellRangeAddress.getFirstColumn()).getStringCellValue());

			mergedCellHeaders.add(sheetHeaderMasterTest);
			headerIndex++;

		}
		return mergedCellHeaders;
	}

	public static void writeExcelTemplate(TemplateMasterTest templateMasterTest)
			throws FileNotFoundException, IOException {
		Set<SheetMasterTest> sheetMasterTestSet = templateMasterTest.getSheetMasterTestList();

		// Converting to list and sorting to get sheets in order (Can be removed if
		// index jumping is allowed while creating sheet)
		List<SheetMasterTest> sheetMasterTestList = sheetMasterTestSet.stream().collect(Collectors.toList());
		Collections.sort(sheetMasterTestList, (o1, o2) -> ((Integer) o1.getSheetIndex()).compareTo(o2.getSheetIndex()));

		try (OutputStream fileOut = new FileOutputStream("test_template.xlsx")) {
			try (XSSFWorkbook workbook = new XSSFWorkbook()) {
				// Adding to handle cell background color issue in poi
				IndexedColorMap colorMap = workbook.getStylesSource().getIndexedColors();

				Iterator<SheetMasterTest> sheetMasterTestListItr = sheetMasterTestList.iterator();
				while (sheetMasterTestListItr.hasNext()) {
					SheetMasterTest sheetMasterTest = sheetMasterTestListItr.next();
					XSSFSheet currentSheet = workbook.createSheet(sheetMasterTest.getSheetName());

					Set<SheetHeaderMasterTest> sheetHeaderMasterTestSet = sheetMasterTest
							.getSheetHeaderMasterTestList();

					for (SheetHeaderMasterTest sheetHeaderMasterTest : sheetHeaderMasterTestSet) {
						int colFrom = sheetHeaderMasterTest.getColFrom();
						int colTo = sheetHeaderMasterTest.getColTo();
						int rowFrom = sheetHeaderMasterTest.getRowFrom();
						int rowTo = sheetHeaderMasterTest.getRowTo();

						XSSFRow row = currentSheet.getRow(rowFrom) == null ? currentSheet.createRow(rowFrom)
								: currentSheet.getRow(rowFrom);
						XSSFCell cell = row.getCell(colFrom) == null ? row.createCell(colFrom) : row.getCell(colFrom);

						cell.setCellValue(sheetHeaderMasterTest.getHeaderText());

						// Enforcing merged region for merged cells
						if ((colFrom != colTo) || (rowFrom != rowTo)) {
							currentSheet.addMergedRegion(new CellRangeAddress(rowFrom, rowTo, colFrom, colTo));
						}

						// Styling to center align headers, can be changed as required
						XSSFCellStyle style = cell.getSheet().getWorkbook().createCellStyle();
						style.setAlignment(HorizontalAlignment.CENTER);
						style.setVerticalAlignment(VerticalAlignment.CENTER);
						XSSFColor headerColor = new XSSFColor(getRGB(sheetHeaderMasterTest.getHeaderColor()),colorMap);
						style.setFillBackgroundColor(headerColor);
						style.setFillPattern(FillPatternType.BIG_SPOTS);
						cell.setCellStyle(style);
					}
				}

				workbook.write(fileOut);
				workbook.close();
			}
			fileOut.flush();
			fileOut.close();

		}

	}

	private static byte[] getRGB(String hexColorString) {

		hexColorString = hexColorString.replace("#", "");
		
		switch (hexColorString.length()) {
		case 6:
			return new byte[] {
					(byte)(Integer.valueOf(hexColorString.substring(0, 2), 16)  & 0xFF),
					(byte)(Integer.valueOf(hexColorString.substring(2, 4), 16)  & 0xFF),
					(byte)(Integer.valueOf(hexColorString.substring(4, 6), 16)  & 0xFF)};
		case 8:
			return new byte[] {
					(byte)(Integer.valueOf(hexColorString.substring(0, 2), 16)  & 0xFF),
					(byte)(Integer.valueOf(hexColorString.substring(2, 4), 16)  & 0xFF),
					(byte)(Integer.valueOf(hexColorString.substring(4, 6), 16)  & 0xFF),
					(byte)(Integer.valueOf(hexColorString.substring(6, 8), 16)  & 0xFF)};
			}
			return null;
		}
	}
