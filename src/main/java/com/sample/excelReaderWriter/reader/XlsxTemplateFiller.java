package com.sample.excelReaderWriter.reader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
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
import org.springframework.beans.BeanUtils;
import org.springframework.util.ResourceUtils;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.sample.excelReaderWriter.entity.SheetHeaderMasterTest;
import com.sample.excelReaderWriter.entity.SheetMasterTest;
import com.sample.excelReaderWriter.entity.TemplateMasterTest;
import com.sample.excelReaderWriter.vo.ExampleVO;

public class XlsxTemplateFiller {

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
									sheetHeaderMasterTest.setHeaderColor(
											cell.getCellStyle().getFillForegroundColorColor().getARGBHex());

									// Setting column data property key
									sheetHeaderMasterTest.setColumnKey(cell.getStringCellValue().toLowerCase());
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

		// Get xlsx header info as list
		List<ExampleVO> allSheetHeaderVOList = createExampleVOList(templateMasterTest);
		// Reproduce template from template entity
		writeExcelTemplate(allSheetHeaderVOList, getJsonData());
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
			sheetHeaderMasterTest.setHeaderColor(
					currentSheet.getRow(cellRangeAddress.getFirstRow()).getCell(cellRangeAddress.getFirstColumn())
					.getCellStyle().getFillForegroundColorColor().getARGBHex());
			// Data of merged cells is always present in the first cell
			sheetHeaderMasterTest.setHeaderText(currentSheet.getRow(cellRangeAddress.getFirstRow())
					.getCell(cellRangeAddress.getFirstColumn()).getStringCellValue());

			// Setting column data property key
			if ((cellRangeAddress.getFirstRow() != cellRangeAddress.getLastRow())
					&& (cellRangeAddress.getFirstColumn() == cellRangeAddress.getLastColumn())) {
				sheetHeaderMasterTest.setColumnKey(currentSheet.getRow(cellRangeAddress.getFirstRow())
						.getCell(cellRangeAddress.getFirstColumn()).getStringCellValue().toLowerCase());
			}

			mergedCellHeaders.add(sheetHeaderMasterTest);
			headerIndex++;

		}
		return mergedCellHeaders;
	}

	public static void writeExcelTemplate(List<ExampleVO> allSheetHeaderVOList, JsonNode parentJson)
			throws FileNotFoundException, IOException {
		Map<Integer, List<ExampleVO>> sheetHeaderMap = allSheetHeaderVOList.stream()
				.collect(Collectors.groupingBy(ExampleVO::getSheetIndex));
		int numOfSheets = sheetHeaderMap.size();
		System.out.println(sheetHeaderMap);

		try (OutputStream fileOut = new FileOutputStream("test_template.xlsx")) {
			try (XSSFWorkbook workbook = new XSSFWorkbook()) {
				// Adding to handle cell background color issue in poi
				IndexedColorMap colorMap = workbook.getStylesSource().getIndexedColors();

				for (int count = 0; count < numOfSheets; count++) {

					List<ExampleVO> sheetHeaderList = sheetHeaderMap.get(count);
					ExampleVO maxRowHeader = Collections.max(sheetHeaderList, Comparator.comparing(ExampleVO::getRowTo));
					int dataRow = maxRowHeader.getRowTo() + 1;
					XSSFSheet currentSheet = workbook.createSheet(sheetHeaderList.get(0).getSheetName());

					// Write header start
					for (ExampleVO exampleVO : sheetHeaderList) {
						int colFrom = exampleVO.getColFrom();
						int colTo = exampleVO.getColTo();
						int rowFrom = exampleVO.getRowFrom();
						int rowTo = exampleVO.getRowTo();

						XSSFRow row = currentSheet.getRow(rowFrom) == null ? currentSheet.createRow(rowFrom)
								: currentSheet.getRow(rowFrom);
						XSSFCell cell = row.getCell(colFrom) == null ? row.createCell(colFrom) : row.getCell(colFrom);

						cell.setCellValue(exampleVO.getHeaderText());

						// Enforcing merged region for merged cells
						if ((colFrom != colTo) || (rowFrom != rowTo)) {
							currentSheet.addMergedRegion(new CellRangeAddress(rowFrom, rowTo, colFrom, colTo));
						}

						// Styling to center align headers, can be changed as required
						XSSFCellStyle style = cell.getSheet().getWorkbook().createCellStyle();
						style.setAlignment(HorizontalAlignment.CENTER);
						style.setVerticalAlignment(VerticalAlignment.CENTER);
						XSSFColor headerColor = new XSSFColor(getRGB(exampleVO.getHeaderColor()), colorMap);
						style.setFillBackgroundColor(headerColor);
						style.setFillPattern(FillPatternType.BIG_SPOTS);
						cell.setCellStyle(style);
					}
					// Write header end
					
					// Create data property name vs xlsx column index map
					Map<String,Integer> propertyColumnMap = sheetHeaderList.stream()
							.filter(o->null!=o.getColumnKey())
							.collect(Collectors.toMap(o->o.getColumnKey(), o->o.getColFrom()));
					// Fetch json data object from parent json for current sheeet - here its sheet name
					ArrayNode dataArrNode = (ArrayNode) parentJson.path(sheetHeaderList.get(0).getSheetName().toLowerCase());
					// Write Data start for current sheet
					Iterator<JsonNode> dataArrItr = dataArrNode.iterator();
					while(dataArrItr.hasNext()) {
						ObjectNode dataObj = (ObjectNode) dataArrItr.next();
						Iterator<String> propertyNames = dataObj.fieldNames();
						XSSFRow row = currentSheet.createRow(dataRow);
						while(propertyNames.hasNext()) {
							String property = propertyNames.next();
							XSSFCell cell = row.createCell(propertyColumnMap.get(property));
							cell.setCellValue(dataObj.path(property).asText());
						}
						
						dataRow++;
					}
					
					// Write header end for current sheet
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
			return new byte[] { (byte) (Integer.valueOf(hexColorString.substring(0, 2), 16) & 0xFF),
					(byte) (Integer.valueOf(hexColorString.substring(2, 4), 16) & 0xFF),
					(byte) (Integer.valueOf(hexColorString.substring(4, 6), 16) & 0xFF) };
		case 8:
			return new byte[] { (byte) (Integer.valueOf(hexColorString.substring(0, 2), 16) & 0xFF),
					(byte) (Integer.valueOf(hexColorString.substring(2, 4), 16) & 0xFF),
					(byte) (Integer.valueOf(hexColorString.substring(4, 6), 16) & 0xFF),
					(byte) (Integer.valueOf(hexColorString.substring(6, 8), 16) & 0xFF) };
		}
		return null;
	}

	private static List<ExampleVO> createExampleVOList(TemplateMasterTest templateMasterTest) {

		List<ExampleVO> exampleVOList = new ArrayList<>();

		List<SheetMasterTest> sheetMasterTestList = new ArrayList<>(templateMasterTest.getSheetMasterTestList());

		for (SheetMasterTest smt : sheetMasterTestList) {

			Set<SheetHeaderMasterTest> ht = smt.getSheetHeaderMasterTestList();
			for (SheetHeaderMasterTest hmt : ht) {
				ExampleVO vo = new ExampleVO();
				BeanUtils.copyProperties(smt, vo);
				BeanUtils.copyProperties(hmt, vo);
				vo.setTemplateId(templateMasterTest.getTemplateId());
				vo.setTemplateName(templateMasterTest.getTemplateName());

				exampleVOList.add(vo);
			}

		}

		return exampleVOList;
	}

	private static JsonNode getJsonData() {
		ObjectMapper mapper = new ObjectMapper();
		ObjectNode parentJson = mapper.createObjectNode();

		// Data for 1st sheet
		ArrayNode personal = mapper.createArrayNode();
		// Data for 1st sheet
		ArrayNode expenses = mapper.createArrayNode();
		
			ObjectNode p1 = mapper.createObjectNode();
			p1.put("name", "SA" + 1);
			p1.put("city", "PB" + 1);
			p1.put("state", "BB" + 1);
			p1.put("pin", "88888" + 1);
			p1.put("pan", "ABCDE1234S" + 1);

			personal.add(p1);

			ObjectNode p2 = mapper.createObjectNode();
			p2.put("name", "SA" + 2);
			p2.put("city", "PB" + 2);
			p2.put("state", "BB" + 2);
			p2.put("pin", "88888" + 2);
			

			personal.add(p2);
			
			ObjectNode p3 = mapper.createObjectNode();
			p3.put("name", "SA" + 3);
			p3.put("city", "PB" + 3);
			
			personal.add(p3);
			
			ObjectNode e1 = mapper.createObjectNode();
			e1.put("date", "2/16/2019" + 1);
			e1.put("amount", "0" + 1);

			expenses.add(e1);
			
			ObjectNode e2 = mapper.createObjectNode();
			e2.put("date", "2/16/2019" + 2);

			expenses.add(e2);
			
			ObjectNode e3 = mapper.createObjectNode();
			e3.put("amount", "0" + 3);

			expenses.add(e3);
		
		parentJson.set("personal", personal);
		parentJson.set("expenses", expenses);
		return parentJson;
	}
}
