package org.wso2.custom;

import org.apache.axiom.om.OMAbstractFactory;
import org.apache.axiom.om.OMElement;
import org.apache.axiom.soap.SOAPBody;
import org.apache.axiom.soap.SOAPEnvelope;
import org.apache.axiom.soap.SOAPFactory;
import org.apache.axis2.AxisFault;
import org.apache.axis2.builder.Builder;
import org.apache.axis2.context.MessageContext;

import javax.xml.namespace.QName;

import java.io.IOException;
import java.io.InputStream;
import java.io.FileNotFoundException;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * Custom axis2 message builder that handle Excel file populate the payload with the csv representation
 */
public class ExcelMessageBuilder implements Builder {

	/**
	 * {@inheritDoc}
	 */
	public OMElement processDocument(InputStream inputStream, String s,
			MessageContext messageContext) throws AxisFault {
		SOAPFactory soapFactory = OMAbstractFactory.getSOAP11Factory();
		SOAPEnvelope soapEnvelope = soapFactory.getDefaultEnvelope();

		try {
			Workbook wb = WorkbookFactory.create(inputStream);
			SOAPBody body = soapEnvelope.getBody();
			
			/* Browse all the sheets of the Excel document */
			for (int i = 0; i < wb.getNumberOfSheets(); i++) {
				/* For each of them create a new XML node containing the csv representation */
				OMElement sheetElement = soapFactory.createOMElement(new QName(
						"sheet"), body);
				sheetElement.addAttribute(soapFactory.createOMAttribute("name",
						null, wb.getSheetAt(i).getSheetName()));
				sheetElement.setText(sheetAsCSV(wb.getSheetAt(i)));
			}

		} catch (FileNotFoundException ex) {
			ex.printStackTrace();
		} catch (IOException ex) {
			ex.printStackTrace();
		} catch (EncryptedDocumentException e) {
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		}

		return soapEnvelope;
	}

	/**
	 * Convert an Excel sheet content in a string containing the csv representation of the sheet.
	 * @param sheet Excel sheet
	 * @return csv representation as a string
	 */
	private String sheetAsCSV(Sheet sheet) {
		String csv = "";
		Row row = null;
		for (int i = 0; i <= sheet.getLastRowNum(); i++) {
			row = sheet.getRow(i);
			String line = "";
			if (row != null) {
				for (int j = 0; j < row.getLastCellNum(); j++) {
					Cell cell = row.getCell(j);
					if (cell != null) {
						switch (cell.getCellType()) {
						case HSSFCell.CELL_TYPE_STRING:
							line += cell.getRichStringCellValue().getString();
							break;
						case HSSFCell.CELL_TYPE_NUMERIC:
							if (HSSFDateUtil.isCellDateFormatted(cell)) {
								/* A date is considered as a numeric type */
								line += getDateFromCell(cell);
							} else {
								line += cell.getNumericCellValue();
							}
							break;
						case HSSFCell.CELL_TYPE_BOOLEAN:
							line += cell.getBooleanCellValue();
						default:
							line += cell;
							break;
						}
					}
					if (j < (row.getLastCellNum() -1)) {
						line += ";";
					}
				}
				if (!isLineEmpty(line)) {
					csv += line + "\n";
				}
			}
		}	
		return csv;
	}

	/**
	 * Method returning the date or time representation of a cell.
	 * @param cell Cell of the Excel document
	 * @return date or time representation of the cell value as a string
	 */
	private String getDateFromCell(Cell cell) {
		Date date = cell.getDateCellValue();
		/* 
		 * TODO : We should find a more generic way of managing TimeZones. 
         * Excel times are enconding in 1889, for some timzone it can be a strange offset
         * like for instance Europe/Paris has an offset of +00:09:21
         * The general idea would be to use the default system time zone and move the date
         * in the current year
         */
		LocalDateTime ldate = LocalDateTime.ofInstant(date.toInstant(), ZoneId.of("CET"));
		DateTimeFormatter format = DateTimeFormatter.ofPattern("yyyy-MM-dd");
		
		if (cell.getNumericCellValue() < 1) {
			format = DateTimeFormatter.ofPattern("HH:mm");
		} 
		return ldate.format(format);
	}
	
	/**
	 * Method checking if a csv line is empty (contains only separator)
	 * @param line csv line as a string
	 * @return true if the line contains no value
	 */
	private boolean isLineEmpty(String line) {
		return line.replaceAll(";", "").trim().isEmpty();
	}
}
