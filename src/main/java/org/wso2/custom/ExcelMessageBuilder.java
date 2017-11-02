package org.wso2.custom;

import org.apache.axiom.om.OMAbstractFactory;
import org.apache.axiom.om.OMElement;
import org.apache.axiom.soap.SOAPEnvelope;
import org.apache.axiom.soap.SOAPFactory;
import org.apache.axis2.AxisFault;
import org.apache.axis2.builder.Builder;
import org.apache.axis2.context.MessageContext;
import org.apache.commons.lang.StringUtils;

import javax.xml.namespace.QName;

import java.io.IOException;
import java.io.InputStream;
import java.io.FileNotFoundException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * Custom axis2 message builder that handle Excel file populate the payload with the csv representation
 */
public class ExcelMessageBuilder implements Builder {
private final static String SEPARATOR = ";";
    
	/**
	 * {@inheritDoc}
	 */
	public OMElement processDocument(InputStream inputStream, String s,
			MessageContext messageContext) throws AxisFault {
		SOAPFactory soapFactory = OMAbstractFactory.getSOAP11Factory();
		SOAPEnvelope soapEnvelope = soapFactory.getDefaultEnvelope();

		try {
			Workbook wb = WorkbookFactory.create(inputStream);
			FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
			
			/* For each of them create a new XML node containing the csv representation */
			for (Sheet sheet : wb) {
				String csv = sheetAsCSV(sheet, evaluator);
				if (StringUtils.isNotEmpty(csv)) {
					OMElement sheetElement = soapFactory.createOMElement(new QName("sheet"), soapEnvelope.getBody());
					sheetElement.addAttribute(soapFactory.createOMAttribute("name",null, sheet.getSheetName()));
					sheetElement.setText(csv);
				}
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
	private String sheetAsCSV(Sheet sheet, FormulaEvaluator evaluator) {
		StringBuilder csv = new StringBuilder();
		StringBuilder lineBuilder = new StringBuilder();
		String  sep, val;
		for (Row row : sheet) {
			if (row != null) {
				sep = "";
				for (Cell cell: row) {
					val = "";
					if (cell != null) {
						switch (evaluator.evaluateInCell(cell).getCellTypeEnum()) {
						case STRING:
							val = cell.getStringCellValue();
							break;
						case NUMERIC:
							if (HSSFDateUtil.isCellDateFormatted(cell)) {
								/* A date is considered as a numeric type */
								val = getDateFromCell(cell);
							} else {
								val = Double.toString(cell.getNumericCellValue());
							}
							break;
						case BOOLEAN:
							val = Boolean.toString(cell.getBooleanCellValue());
						case BLANK:
							val = "";
						default:
							val = cell.toString();
							break;
						}
					}
					lineBuilder.append(sep).append(val);
					sep = SEPARATOR;
				}
				if (!isLineEmpty(lineBuilder.toString())) {
					csv.append(lineBuilder).append(System.getProperty("line.separator"));
				}
			}
			lineBuilder.setLength(0);
		}	
		return csv.toString();
	}

	/**
	 * Method returning the date or time representation of a cell.
	 * @param cell Cell of the Excel document
	 * @return date or time representation of the cell value as a string
	 */
	private String getDateFromCell(Cell cell) {
		Date date = cell.getDateCellValue();
		String dateFormatString = cell.getCellStyle().getDataFormatString();

		/* 
		* TODO : We should find a more generic way of managing TimeZones. 
         	* Excel times are enconding in 1889, for some timzone it can be a strange offset
         	* like for instance Europe/Paris has an offset of +00:09:21
		* The general idea would be to use the default system time zone and move the date
         	* in the current year
         	*/
		SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
		if (cell.getNumericCellValue() <= 1) {
			/* Date till 1900-01-01 23:59 are considered as time */
			format = new SimpleDateFormat("HH:mm");
		}  else if (dateFormatString.toLowerCase().contains("y") && dateFormatString.toLowerCase().contains("h")) {
			/* Case of date and time (Based on Excel cell format) */
			format = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss");
		}
		return format.format(date);
	}
	
	/**
	 * Method checking if a csv line is empty (contains only separator)
	 * @param line csv line as a string
	 * @return true if the line contains no value
	 */
	private boolean isLineEmpty(String line) {
		return line.replaceAll(SEPARATOR, "").trim().isEmpty();
	}
}
