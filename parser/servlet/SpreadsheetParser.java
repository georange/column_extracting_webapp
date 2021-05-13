package com.sheret.spreadsheet.parser.servlet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;
import java.io.Reader;
import java.io.Writer;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.ServletException;
import javax.servlet.annotation.MultipartConfig;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.Part;
import javax.swing.table.DefaultTableModel;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVPrinter;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.jdom.input.SAXBuilder;
import org.jopendocument.dom.spreadsheet.SpreadSheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Attr;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import com.monitorjbl.xlsx.StreamingReader;
import com.sheret.spreadsheet.parser.serialize.XMLToXPathIniSerializer;
import com.sheret.spreadsheet.parser.utils.HttpHeaderUtils;

import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

/**********************************************************/
/**********************************************************/
@WebServlet("/SpreadsheetParser")
@MultipartConfig(location = "/tmp", fileSizeThreshold = 1024 * 1024, maxFileSize = 1024 * 1024
		* 100, maxRequestSize = 1024 * 1024 * 100 * 5)

public class SpreadsheetParser extends HttpServlet {
	private static Logger logger = LoggerFactory.getLogger(SpreadsheetParser.class);

	private static final long serialVersionUID = 1L;

	/**********************************************************/
	/**********************************************************/
	public SpreadsheetParser() {
		super();
	}

	protected void doGet(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		logger.debug("Here");

		String outputType = request.getParameter("output");

		response.setStatus(200);
		if ((outputType != null) && (outputType.equalsIgnoreCase("text"))) {
			response.setContentType("text/plain");
			response.getWriter().println("This is Georgia's first servlet.");
			response.getWriter().println("This is is plain text.");
		} else {
			response.setContentType("text/html");
			response.getWriter().println(
					"<html><head><title>Title</title></head><body><h1>This is Georgia's first servlet.</h1>This is HTML.</body></html>");
		}

		// printRequest(request, response);

		return;
	}

	/**********************************************************/
	/**********************************************************/
	protected void doPost(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		Part filePart = null;
		String fileContentType = null;
		
		String filename = null;
		String outputFormat = null;
		String[] columns = null;
		boolean quotes = true;
		
		String outfile;
		File tempFile;
		InputStream inStream = null;
		FileOutputStream fileOutStream;
		long fileSize = 0;

		String mainContextType = request.getContentType();

		if ((mainContextType != null) && (mainContextType.equals("multipart/form-data"))) {
			// get file and filename
			filePart = request.getPart("file");
		}
		
		if (filePart != null) {
			logger.debug("Found file part.");
			inStream = filePart.getInputStream();

			//fileContentType = filePart.getContentType();
			fileSize = filePart.getSize();
			//logger.debug("fileContentType: " + fileContentType);
			logger.debug("filePart size: " + fileSize);
			//response.getWriter().println(fileContentType);

			try {
				filename = HttpHeaderUtils.getFilenameFromPart(filePart);
				logger.debug("filename: " + filename);

				// fileContentType = filename.substring(filename.lastIndexOf('.')+1);

			} catch (Exception e) {
				response.setStatus(400);
				logger.debug("Failed to get filename from part: " + e.toString());
				e.printStackTrace();
			}
		} else {
			logger.debug("Did NOT find file part.");
			inStream = request.getInputStream();
			//fileContentType = request.getContentType();
			fileSize = request.getContentLength();
			//logger.debug("fileContentType: " + fileContentType);
			logger.debug("filePart size: " + fileSize);
			//response.getWriter().println(fileContentType);
			try {
				filename = HttpHeaderUtils.getFilenameFromHeaderValue(request.getHeader("Content-Disposition"));
				logger.debug("filename: " + filename);
				// fileContentType = filename.substring(filename.lastIndexOf('.')+1);

			} catch (Exception e) {
				response.setStatus(400);
				logger.debug("Failed to get filename from header: " + e.toString());
				e.printStackTrace();
			}
		}

		//printRequest(request, response);
		
		try {
			// check that filename and fileContentType exist			
			if (filename == null) {
				throw new MissingParamException("Missing Parameter: filename\n");
			} 
			if (fileSize == 0) {
				throw new MissingParamException("Missing Parameter: fileSize\n");
			}
			
			fileContentType = filename.substring(filename.lastIndexOf(".")+1);
			logger.debug("fileContentType: " + fileContentType);
			if (fileContentType == null) {
				throw new MissingParamException("Missing Parameter: fileContentType\n");
			}
			
			// get output format
			outputFormat = request.getParameter("outputFormat");
			logger.debug("Found outputFormat: " + outputFormat);
			if (outputFormat == null) {
				throw new MissingParamException("Missing Parameter: outputFormat\n");
			}
	
			// get columns as array
			String columnString = request.getParameter("columns");
			logger.debug("Found columnString: " + columnString);
			if (columnString == null) {
				throw new MissingParamException("Missing Parameter: columnString\n");
			}
			columns = columnString.split("\\s*,\\s*");
	
			// get quotes
			String quotesString = request.getParameter("quotes");
			if (quotesString == null) {
				throw new MissingParamException("Missing Parameter: quotesString\n");
			}
			if (quotesString.equals("true")) {
				quotes = true;
			} else {
				quotes = false;
			}
			logger.debug("Found quotes: " + quotes);
			
			// copy file to temp file
			if (inStream != null) {
				tempFile = File.createTempFile("tempfile", ".txt");
				tempFile.deleteOnExit();

				fileOutStream = new FileOutputStream(tempFile);

				IOUtils.copy(inStream, fileOutStream);
				fileOutStream.flush();
				fileOutStream.close();

				// TODO: here's where you change where the file outputs to
				// set up outfile name by modifying filename
				//outfile = "/Users/a9953/Desktop/" + filename.substring(0, filename.lastIndexOf('.')).replaceAll(" ", "_") + "parsed";
				outfile = "/home/CommonDocs/.JTWinShare/" + filename.substring(0, filename.lastIndexOf('.')).replaceAll(" ", "_") + "parsed";

				doConversion(response, outfile, outputFormat, filename, fileContentType, columns, quotes, tempFile);
			}

		} catch (MissingParamException e) {
			response.setStatus(400);
			response.getWriter().println(e.getMessage());
			logger.debug(e.getMessage());
			e.printStackTrace();
		}
		return;
	}

	// does the converting
	public void doConversion(HttpServletResponse response, String outfile, String outputFormat, String filename,
			String fileContentType, String[] columns, boolean quotes, File tempFile) throws IOException {
		try {
			// for converting CSVs to everything else
			if (fileContentType.equals("csv")) {
				if (outputFormat.equals("csv")) {
					outfile = outfile + ".csv";
					response.getWriter().println(outfile + "\n");
					CSVtoCSV(tempFile.getAbsolutePath(), outfile, columns, response, quotes);

				} else if (outputFormat.equals("xls")) {
					outfile = outfile + ".xls";
					response.getWriter().println(outfile + "\n");
					CSVtoXLS(tempFile.getAbsolutePath(), outfile, columns, response, quotes);

				} else if (outputFormat.equals("xlsx")) {
					outfile = outfile + ".xlsx";
					response.getWriter().println(outfile + "\n");
					CSVtoXLSX(tempFile.getAbsolutePath(), outfile, columns, response, quotes);

				} else if (outputFormat.equals("xml")) {
					outfile = outfile + ".xml";
					response.getWriter().println(outfile + "\n");
					CSVtoXML(tempFile.getAbsolutePath(), outfile, columns, response, quotes);

				} else if (outputFormat.equals("ods")) {
					outfile = outfile + ".ods";
					response.getWriter().println(outfile + "\n");
					CSVtoODS(tempFile.getAbsolutePath(), outfile, columns, response, quotes);

				} else if (outputFormat.equals("xini")) {
					outfile = outfile + ".xini";
					response.getWriter().println(outfile + "\n");
					CSVtoXINI(tempFile.getAbsolutePath(), outfile, columns, response, quotes);
				} else {
					throw new ConversionNotFoundException("Error: Could not find conversion: " 
							+ fileContentType + " to " + outputFormat + "\n");
				}

				// for outputting to CSV from everything else
			} else if (outputFormat.equals("csv")) {
				outfile = outfile + ".csv";
				response.getWriter().println(outfile + "\n");

				if (fileContentType.equals("xls")) {
					XLStoCSV(tempFile.getAbsolutePath(), outfile, columns, response, quotes);
				} else if (fileContentType.equals("xlsx")) {
					XLSXtoCSV(tempFile.getAbsolutePath(), outfile, columns, response, quotes);
				} else if (fileContentType.equals("xml")) {
					XMLtoCSV(tempFile.getAbsolutePath(), outfile, columns, response, quotes);
				} else if (fileContentType.equals("ods")) {
					ODStoCSV(tempFile.getAbsolutePath(), outfile, columns, response, quotes);
				} else {
					throw new ConversionNotFoundException("Error: Could not find conversion: " 
							+ fileContentType + " to " + outputFormat + "\n");
				}

			} else {
				// XML to XINI
				if (fileContentType.equals("xml") && outputFormat.equals("xini")) {
					outfile = outfile + ".xini";
					response.getWriter().println("Note that XML to XINI will not consider columns or quote options.");
					response.getWriter()
							.println("If those are of concern, please convert from a different file type.\n");
					XMLtoXINI(tempFile, outfile, response);

					// all other cases
				} else {
					// make a temporary csv file
					String tempCsvName = outfile.substring(outfile.lastIndexOf('/') + 1).replaceAll(" ", "_") + ".csv";
					File tempCsvFile = tempCSV(tempCsvName, response);

					if (fileContentType.equals("xls")) {
						XLStoCSV(tempFile.getAbsolutePath(), tempCsvFile.getAbsolutePath(), columns, response, quotes);
					} else if (fileContentType.equals("xlsx")) {
						XLSXtoCSV(tempFile.getAbsolutePath(), tempCsvFile.getAbsolutePath(), columns, response, quotes);
					} else if (fileContentType.equals("ods")) {
						ODStoCSV(tempFile.getAbsolutePath(), tempCsvFile.getAbsolutePath(), columns, response, quotes);
					} else {
						tempCsvFile.delete();
						tempFile.delete();
						throw new ConversionNotFoundException("Error: Could not find conversion: " 
								+ fileContentType + " to " + outputFormat + "\n");
					}
					// then run through the other side
					if (outputFormat.equals("xml")) {
						outfile = outfile + ".xml";
						CSVtoXML(tempCsvFile.getAbsolutePath(), outfile, columns, response, quotes);
					} else if (outputFormat.equals("xini")) {
						outfile = outfile + ".xini";
						CSVtoXINI(tempCsvFile.getAbsolutePath(), outfile, columns, response, quotes);
					} else if (outputFormat.equals("xls")) {
						outfile = outfile + ".xls";
						CSVtoXLS(tempCsvFile.getAbsolutePath(), outfile, columns, response, quotes);
					} else if (outputFormat.equals("xlsx")) {
						outfile = outfile + ".xlsx";
						CSVtoXLSX(tempCsvFile.getAbsolutePath(), outfile, columns, response, quotes);
					} else if (outputFormat.equals("ods")) {
						outfile = outfile + ".ods";
						CSVtoODS(tempCsvFile.getAbsolutePath(), outfile, columns, response, quotes);
					} else {
						throw new ConversionNotFoundException("Error: Could not find conversion: " 
								+ fileContentType + " to " + outputFormat + "\n");
					}

					tempCsvFile.delete();
				}

			}
		} catch (MissingParamException e) {
			response.setStatus(400);
			response.getWriter().println(e.getMessage());
			logger.debug(e.getMessage());
			e.printStackTrace();
		} catch (Exception e) {
			logger.debug("Conversion fail: " + e.toString());
			response.setStatus(400);
			e.printStackTrace();
		}
		// clean up and leave
		logger.debug("Temp file is: " + tempFile.getAbsolutePath());
		tempFile.delete();

		return;
	}

	public void printRequest(HttpServletRequest req, HttpServletResponse res) throws IOException {

		res.setContentType("text/plain");

		Enumeration<String> headerNames = req.getHeaderNames();

		while (headerNames.hasMoreElements()) {

			String headerName = headerNames.nextElement();
			res.getWriter().println(headerName);
			res.getWriter().println("\n");

			Enumeration<String> headers = req.getHeaders(headerName);
			while (headers.hasMoreElements()) {
				String headerValue = headers.nextElement();
				res.getWriter().println("\t" + headerValue);
				res.getWriter().println("\n");
			}

		}
	}

	private static void CSVtoCSV(String filePath, String outfile, String[] cols, HttpServletResponse response,
			boolean quotes) {
		Reader reader;
		Writer writer;
		CSVPrinter csvPrinter;

		try {
			reader = Files.newBufferedReader(Paths.get(filePath));
			writer = new FileWriter(outfile);
			CSVParser csvParser = new CSVParser(reader,
					CSVFormat.DEFAULT.withFirstRecordAsHeader().withIgnoreHeaderCase().withTrim());
			csvPrinter = getCsvPrinter(writer, cols, quotes, response);

			if (csvPrinter == null) {
				logger.debug("CSV to CSV fail. CSVPrinter not made.");
			}

			for (CSVRecord csvRecord : csvParser) {
				String[] currCol = new String[cols.length];
				for (int i = 0; i < cols.length; i++) {
					currCol[i] = csvRecord.get(cols[i]);
					if (quotes == false) {
						currCol[i] = currCol[i].replaceAll("^\"|\"$", "");
					}
				}
				csvPrinter.printRecord(Arrays.asList(currCol));

				// testing output
				for (String item : currCol) {
					response.getWriter().print(item + " ");

				}
				response.getWriter().println();

				csvPrinter.flush();
			}
			csvParser.close();
			csvPrinter.close();

		} catch (Exception e) {
			response.setStatus(400);
			logger.debug("CSV to CSV fail. Exception: " + e.toString());
			e.printStackTrace();
		}

		return;
	}
	
	private static void CSVtoXLSX(String filePath, String outfile, String[] cols, HttpServletResponse response,
			boolean quotes) {
		Reader reader;
		SXSSFWorkbook workbook = new SXSSFWorkbook(100); // keep 100 rows in memory, exceeding rows will be flushed to disk
        Sheet sheet = workbook.createSheet(outfile.substring(outfile.lastIndexOf('/') + 1, outfile.lastIndexOf('.')));
		// create a header Row
		Row headerRow = sheet.createRow(0);

		// create header cells
		for (int i = 0; i < cols.length; i++) {
			Cell cell = headerRow.createCell(i);
			if (quotes == true) {
				cell.setCellValue(cols[i]);
			} else {
				cell.setCellValue(cols[i].replaceAll("^\"|\"$", ""));
			}
			cell.setCellStyle(workbook.getCellStyleAt(i));
		}
		
		try {
			reader = Files.newBufferedReader(Paths.get(filePath));
			CSVParser csvParser = new CSVParser(reader,
					CSVFormat.DEFAULT.withFirstRecordAsHeader().withIgnoreHeaderCase().withTrim());

			int rowNum = 1;
			for (CSVRecord csvRecord : csvParser) {
				Row row = sheet.createRow(rowNum++);
				for (int i = 0; i < cols.length; i++) {
					String item = csvRecord.get(cols[i]);
					if (quotes == false) {
						item = item.replaceAll("^\"|\"$", "");
					}
					row.createCell(i).setCellValue(item);
				}
			}
//			// resize all columns to fit the content size
//			for (int i = 0; i < cols.length; i++) {
//				sheet.autoSizeColumn(i);
//			}

			// write the output to a file
			FileOutputStream fileOut = new FileOutputStream(outfile);
			workbook.write(fileOut);

			fileOut.close();
			csvParser.close();
			workbook.close();

		} catch (Exception e) {
			response.setStatus(400);
			logger.debug("CSV to XLS(X) fail. Exception: " + e.toString());
			e.printStackTrace();
		}

		return;
	}

	private static void CSVtoXLS(String filePath, String outfile, String[] cols, HttpServletResponse response,
			boolean quotes) {
		Reader reader;
		WritableWorkbook workbook;
		
		try {
			reader = Files.newBufferedReader(Paths.get(filePath));
			CSVParser csvParser = new CSVParser(reader,
					CSVFormat.DEFAULT.withFirstRecordAsHeader().withIgnoreHeaderCase().withTrim());
			workbook = jxl.Workbook.createWorkbook(new File(outfile));
		
			// create a Sheet
			WritableSheet sheet = workbook
					.createSheet(outfile.substring(outfile.lastIndexOf('/') + 1, outfile.lastIndexOf('.')), 0);
		
			// create header row
			Label cell = null;
			for (int i = 0; i < cols.length; i++) {
				// Label cell = headerRow.createCell(i);
				if (quotes == true) {
					cell = new Label(i, 0, cols[i]);
				} else {
					cell = new Label(i, 0, cols[i].replaceAll("^\"|\"$", ""));
				}
				sheet.addCell(cell);
			}
			
			// add the rest of the rows
			int rowNum = 1;
			for (CSVRecord csvRecord : csvParser) {
				for (int i = 0; i < cols.length; i++) {
					String item = csvRecord.get(cols[i]);

					if (quotes == false) {
						item = item.replaceAll("^\"|\"$", "");
					}
					// TODO: max rows exceeded???? could push onto more sheets but why..
					cell = new Label(i, rowNum, item);
					sheet.addCell(cell);
				}
				rowNum++;
			}

			// write the output and clean up
			workbook.write();
			csvParser.close();
			workbook.close();

		} catch (Exception e) {
			response.setStatus(400);
			logger.debug("CSV to XLS(X) fail. Exception: " + e.toString());
			e.printStackTrace();
		}

		return;
	}

	private static void CSVtoXML(String filePath, String outfile, String[] cols, HttpServletResponse response,
			boolean quotes) {
		Reader reader;

		try {
			reader = Files.newBufferedReader(Paths.get(filePath));
			CSVParser csvParser = new CSVParser(reader,
					CSVFormat.DEFAULT.withFirstRecordAsHeader().withIgnoreHeaderCase().withTrim());

			DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder docBuilder = docFactory.newDocumentBuilder();

			// root elements
			Document doc = docBuilder.newDocument();
			Element rootElement = doc.createElement(
					outfile.substring(outfile.lastIndexOf('/') + 1, outfile.lastIndexOf('.') - 6).replaceAll(" ", "_"));
			doc.appendChild(rootElement);

			for (CSVRecord csvRecord : csvParser) {
				// row elements
				Element row = doc.createElement("entry");
				rootElement.appendChild(row);

				// set attribute entry element
				Attr attr = doc.createAttribute("num");
				attr.setValue("" + csvRecord.getRecordNumber());
				row.setAttributeNode(attr);

				for (int i = 0; i < cols.length; i++) {
					Element curr;
					if (quotes == true) {
						curr = doc.createElement(cols[i].replaceAll("^\"|\"$", "").replaceAll(" ", "_"));
						curr.appendChild(doc.createTextNode(csvRecord.get(cols[i])));
					} else {
						curr = doc.createElement(cols[i].replaceAll("^\"|\"$", "").replaceAll(" ", "_"));
						curr.appendChild(doc.createTextNode(csvRecord.get(cols[i]).replaceAll("^\"|\"$", "")));
					}

					row.appendChild(curr);
				}

				// write the content into xml file
				TransformerFactory transformerFactory = TransformerFactory.newInstance();
				Transformer transformer = transformerFactory.newTransformer();
				transformer.setOutputProperty(OutputKeys.INDENT, "yes");
				transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "4");
				transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "yes");

				DOMSource source = new DOMSource(doc);
				StreamResult result = new StreamResult(new File(outfile));

				transformer.transform(source, result);

			}
			csvParser.close();

		} catch (Exception e) {
			response.setStatus(400);
			logger.debug("CSV to XML fail. Exception: " + e.toString());
			e.printStackTrace();
		}

		return;
	}

	private static void CSVtoODS(String filePath, String outfile, String[] cols, HttpServletResponse response,
			boolean quotes) {
		Reader reader;
		DefaultTableModel table;

		try {
			reader = Files.newBufferedReader(Paths.get(filePath));
			CSVParser csvParser = new CSVParser(reader,
					CSVFormat.DEFAULT.withFirstRecordAsHeader().withIgnoreHeaderCase().withTrim());

			String[] tempCols = new String[cols.length];
			if (quotes == false) {
				for (int i = 0; i < cols.length; i++) {
					tempCols[i] = cols[i].replaceAll("^\"|\"$", "");
				}
				table = new DefaultTableModel(tempCols, 0);
			} else {
				table = new DefaultTableModel(cols, 0);
			}

			for (CSVRecord csvRecord : csvParser) {
				String[] currRow = new String[cols.length];
				for (int i = 0; i < cols.length; i++) {
					currRow[i] = csvRecord.get(cols[i]);
					if (quotes == false) {
						currRow[i] = currRow[i].replaceAll("^\"|\"$", "");
					}
				}
				table.addRow(currRow);
			}
			csvParser.close();

			// save the data to an ODS file
			final File file = new File(outfile);
			SpreadSheet.createEmpty(table).saveAs(file);

		} catch (Exception e) {
			response.setStatus(400);
			logger.debug("CSV to ODS fail. Exception: " + e.toString());
			e.printStackTrace();
		}

		return;
	}

	private static void CSVtoXINI(String filePath, String outfile, String[] cols, HttpServletResponse response,
			boolean quotes) {
		File tempFile;
		String tempName = outfile.substring(outfile.lastIndexOf('/') + 1, outfile.lastIndexOf('.')).replaceAll(" ", "_")
				+ ".xml";

		try {
			// temporary in-between xml file
			tempFile = new File(tempName);
			tempFile.deleteOnExit();
			logger.debug("Temp xml file is: " + tempFile.getAbsolutePath());
			response.getWriter().println(tempName + "\n");

			CSVtoXML(filePath, tempFile.getAbsolutePath(), cols, response, quotes);

			XMLtoXINI(tempFile, outfile, response);

			tempFile.delete();
		} catch (IOException e) {
			response.setStatus(400);
			logger.debug("CSV to XINI fail. Exception: " + e.toString());
			e.printStackTrace();
		} catch (Exception e) {
			response.setStatus(400);
			logger.debug("CSV to XINI fail. Exception: " + e.toString());
			e.printStackTrace();
		}

		return;
	}
	
	private static void XLSXtoCSV(String filePath, String outfile, String[] cols, HttpServletResponse response,
			boolean quotes) throws IOException {
		Writer writer;
		CSVPrinter csvPrinter;
		try {
			writer = new FileWriter(outfile);
			csvPrinter = getCsvPrinter(writer, cols, quotes, response);

			if (csvPrinter == null) {
				logger.debug("XLSX to CSV fail. CSVPrinter not made.");
			}
			
	
			InputStream input = new FileInputStream(new File(filePath));
			Workbook workbook = StreamingReader.builder()
		        .rowCacheSize(100)    // number of rows to keep in memory (defaults to 10)
		        .bufferSize(4096)     // buffer size to use when reading InputStream to file (defaults to 1024)
		        .open(input);         // InputStream or File for XLSX file (required)
		
			boolean sheetFound = false;
			int limit = 5;
			int count = 0;
			for (Sheet sheet : workbook) {
				//System.out.println(sheet.getSheetName());
				Map<String, Integer> colMapByName = null;
				for (Row row : sheet) {
					if (count == limit) {
						count = 0;
						break;
					}
					// first row of non-headers
					if (sheetFound == true) {
						// copy wanted columns into each row of the output
						Cell cell;
						String[] currCol = new String[cols.length];
						for (int j = 0; j < cols.length; j++) {

							cell = row.getCell(colMapByName.get(cols[j]));
							currCol[j] = getXlCellVal(cell);

							if (quotes == false) {
								currCol[j] = currCol[j].replaceAll("^\"|\"$", "");
							}
						}
						csvPrinter.printRecord(Arrays.asList(currCol));
						csvPrinter.flush();
						continue;
					} else {
						// look for the header that matches the columns we are looking for
						colMapByName = new HashMap<String, Integer>();
						int i = 0;
						for (Cell cell : row) {
							// create mapping here
							colMapByName.put(getXlCellVal(cell), i);
							i++;
							//System.out.println(cell.getStringCellValue());
						}
						// check if found here
						if (hasAllCols(colMapByName, cols) == true) {
							sheetFound = true;
						} else {
							count++;
						}
					}
				}
				// will not look for instances of same headers in later sheets 
				if (sheetFound == true) {
					break;
				}
			}
			if (sheetFound == false) {
				throw new ColsNotFoundException("Conversion fail: specified columns not found.\n");
			}
		
		} catch (ColsNotFoundException e) {
			response.setStatus(400);
			response.getWriter().println(e.getMessage());
			logger.debug(e.getMessage());
			e.printStackTrace();
		} catch (Exception e) {
			response.setStatus(400);
			logger.debug("XLAS to CSV fail. Exception: " + e.toString());
			e.printStackTrace();
		}
		return;		
	}

	// works for xls and xlsx
	private static void XLStoCSV(String filePath, String outfile, String[] cols, HttpServletResponse response,
			boolean quotes) throws IOException {
		Writer writer;
		CSVPrinter csvPrinter;
		try {
			writer = new FileWriter(outfile);
			csvPrinter = getCsvPrinter(writer, cols, quotes, response);

			if (csvPrinter == null) {
				logger.debug("XL to CSV fail. CSVPrinter not made.");
			}
			
			// creating a Workbook from an Excel file 
			jxl.Workbook workbook = jxl.Workbook.getWorkbook(new File(filePath));

			// retrieving the number of sheets in the Workbook
			int numberOfSheets = workbook.getNumberOfSheets();
			logger.debug("Sheets found:" + numberOfSheets);
			jxl.Cell cell;
			jxl.Sheet sheet;
			boolean sheetFound = false;
			
			// for each sheet
			for (int i = 0; i < numberOfSheets; i++) {
				sheet = workbook.getSheet(i);
				Map<String, Integer> colMapByName = null;
				
				// check for matching headers in the first 5 rows
				int limit = 5;
				int headerRow = 0;
				
				for (int row = 0; row < limit; row++) {
					// create a mapping between column names and indices
					colMapByName = new HashMap<String, Integer>();
					jxl.Cell[] cells = sheet.getRow(row);
					
					for (int j = 0; j < cells.length; j++) {
						cell = cells[j];
						colMapByName.put(cell.getContents(), j);
						//logger.debug(cell.getContents() + " : " + j);
					}
					// if all cols found, break the row loop,
					// otherwise continue to next sheet
					if (hasAllCols(colMapByName, cols) == true) {
						sheetFound = true;
						headerRow = row;
						break;
					}
				}
				if (sheetFound == true) {
					String[] currCol = new String[cols.length];
					int rowNum = sheet.getRows();
					//logger.debug("length of cols: " + cols.length);
					
					for (int j = headerRow+1; j < rowNum; j++) {
						for (int k = 0; k < cols.length; k++) {
							int index = colMapByName.get(cols[k]);
							//logger.debug("row indexes?: " + j);
							cell = sheet.getCell(index, j);
							currCol[k] = cell.getContents();

							if (quotes == false) {
								currCol[k] = currCol[k].replaceAll("^\"|\"$", "");
							}
						}
						csvPrinter.printRecord(Arrays.asList(currCol));
						csvPrinter.flush();
					}
					break;
				}
			}
			if (sheetFound == false) {
				throw new ColsNotFoundException("Conversion fail: specified columns not found.\n");
			}
			
			csvPrinter.close();
			workbook.close();
		} catch (ColsNotFoundException e) {
			response.setStatus(400);
			response.getWriter().println(e.getMessage());
			logger.debug(e.getMessage());
			e.printStackTrace();
		} catch (Exception e) {
			response.setStatus(400);
			logger.debug("XL to CSV fail. Exception: " + e.toString());
			e.printStackTrace();
		}
		return;
	}
	
	private static boolean hasAllCols(Map<String, Integer> map, String[] cols) {
		for (int i = 0; i < cols.length; i++) {
			if (map.get(cols[i]) == null) {
				return false;
			}
		}
		return true;
	}

	// helper for XLtoCSV
	private static String getXlCellVal(Cell cell) {
		switch (cell.getCellType()) {
		case BOOLEAN:
			return cell.getBooleanCellValue() + "";
		case STRING:
			return cell.getRichStringCellValue().getString();
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
				return sdf.format(DateUtil.getJavaDate(DateUtil.getExcelDate(cell.getDateCellValue())));

			} else {
				// How to tell between currency or regular number? can't...
				// NumberFormat formatter = NumberFormat.getCurrencyInstance();
				// return formatter.format(cell.getNumericCellValue()).trim();
				return cell.getNumericCellValue() + "";
			}
		case FORMULA:
			return cell.getCellFormula();
		case BLANK:
			return "";
		default:
			return "";
		}
	}

	// This only works on XMLs created by this program..
	private static void XMLtoCSV(String filePath, String outfile, String[] cols, HttpServletResponse response,
			boolean quotes) {
		Writer writer;
		CSVPrinter csvPrinter;
		File inputFile = new File(filePath);
		SAXBuilder saxBuilder = new SAXBuilder();

		try {
			writer = new FileWriter(outfile);
			csvPrinter = getCsvPrinter(writer, cols, quotes, response);

			if (csvPrinter == null) {
				logger.debug("XML to CSV fail. CSVPrinter not made.");
			}

			org.jdom.Document document = saxBuilder.build(inputFile);
			org.jdom.Element rootElement = document.getRootElement();

			// any way to avoid suppression ????
			@SuppressWarnings("unchecked")
			List<org.jdom.Element> nodeList = rootElement.getChildren();

			for (int i = 0; i < nodeList.size(); i++) {
				org.jdom.Element item = nodeList.get(i);
				String[] currCol = new String[cols.length];

				for (int j = 0; j < cols.length; j++) {
					currCol[j] = item.getChild(cols[j].replaceAll(" ", "_")).getText();
					if (quotes == false) {
						currCol[j] = currCol[j].replaceAll("^\"|\"$", "");
					}
				}
				csvPrinter.printRecord(Arrays.asList(currCol));
				csvPrinter.flush();
			}
			csvPrinter.close();

		} catch (Exception e) {
			response.setStatus(400);
			logger.debug("XML to CSV fail. Exception: " + e.toString());
			e.printStackTrace();
		}

		return;
	}

	private static CSVPrinter getCsvPrinter(Writer writer, String[] cols, boolean quotes, HttpServletResponse response) {
		CSVPrinter csvPrinter = null;

		String[] tempCols = new String[cols.length];

		try {
			if (quotes == false) {
				for (int i = 0; i < cols.length; i++) {
					tempCols[i] = cols[i].replaceAll("^\"|\"$", "");
				}
				csvPrinter = new CSVPrinter(writer, CSVFormat.DEFAULT.withHeader(tempCols));// .withQuoteMode(QuoteMode.MINIMAL));
				return csvPrinter;
			} else {
				csvPrinter = new CSVPrinter(writer, CSVFormat.DEFAULT.withHeader(cols));// .withQuoteMode(QuoteMode.MINIMAL));
				return csvPrinter;
			}
		} catch (Exception e) {
			response.setStatus(400);
			logger.debug("getCsvPrinter fail. Exception: " + e.toString());
			e.printStackTrace();
		}
		return csvPrinter;
	}

	private static void ODStoCSV(String filePath, String outfile, String[] cols, HttpServletResponse response,
			boolean quotes) throws IOException {
		Writer writer;
		CSVPrinter csvPrinter;
		File inputFile = new File(filePath);
		SpreadSheet spreadsheet;

		try {
			writer = new FileWriter(outfile);
			csvPrinter = getCsvPrinter(writer, cols, quotes, response);

			if (csvPrinter == null) {
				logger.debug("ODS to CSV fail. CSVPrinter not made.");
			}

			spreadsheet = SpreadSheet.createFromFile(inputFile);
			int numberOfSheets = spreadsheet.getSheetCount();
			logger.debug("Sheets found:" + numberOfSheets);
			org.jopendocument.dom.spreadsheet.Sheet sheet;
			org.jopendocument.dom.spreadsheet.Cell<SpreadSheet> cell;

			boolean sheetFound = false;
			for (int i = 0; i < numberOfSheets; i++) {
				sheet = spreadsheet.getSheet(i);

				// stop after a certain number of consecutive empty columns
				int count = 0;

				// check for matching columns in first few rows
				int limit = 5;
				int headerRow = 0;
				
				// create a mapping between column names and indices
				Map<String, Integer> colMapByName = null; 
				for (int row = 0; row < limit; row++) {
					colMapByName = new HashMap<String, Integer>();
					
					for (int col = 0; col < sheet.getColumnCount(); col++) {
						// map the value of each column in current row to the corresponding column index
						cell = sheet.getImmutableCellAt(col, row);
						String temp = sheet.getImmutableCellAt(col, row).getTextValue();
						if ("".equals(temp) == false) {
							String cellText = cell.getTextValue().trim();
							colMapByName.put(cellText, col);
							count = 0;
						} else {
							count++;
							if (count == 4) {
								break;
							}
						}
						// account for spanning cells?
						// col = col + cell.getColumnsSpanned() - 1;
					}
					// if all cols found, break the row loop,
					// if last row of last sheet throw exception, 
					// otherwise continue to next sheet
					if (hasAllCols(colMapByName, cols) == true) {
						sheetFound = true;
						headerRow = row;
						break;
					} else if (row == limit - 1 && i == numberOfSheets - 1 && sheetFound == false) {
						throw new ColsNotFoundException("Conversion fail: specified columns not found.\n");
					}
					count = 0;
				}
				if (sheetFound == true) {
					// stop after a certain number of consecutive empty rows
					boolean rowEmpty = true;
					count = 0;

					// iterate over rows
					for (int row = headerRow + 1; row < sheet.getRowCount(); row++) {
						// go through columns
						String[] currCol = new String[cols.length];
						for (int j = 0; j < cols.length; j++) {

							if (colMapByName.get(cols[j]) != null) {
								cell = sheet.getImmutableCellAt(colMapByName.get(cols[j]), row);
								String temp = getOdsCellVal(cell);

								if ("".equals(temp) == false) {
									currCol[j] = temp;
									rowEmpty = false;

									if (quotes == false) {
										currCol[j] = currCol[j].replaceAll("^\"|\"$", "");
									}
								}
								// make sure the next column hasn't been taken
								// by a
								// span
								// if (j + 1 < cols.length) {
								// if (colMapByName.get(cols[j]) >
								// colMapByName.get(cols[j + 1])) {
								// while (j + cell.getColumnsSpanned() >
								// colMapByName.get(cols[j + 1])) {
								// j = j + 1;
								// }
								// }
								// }
							}

						}
						if (rowEmpty == false) {
							csvPrinter.printRecord(Arrays.asList(currCol));
							csvPrinter.flush();
							rowEmpty = true;
							count = 0;

						} else {
							// count empty rows and break when a limit is
							// reached
							count++;
							if (count == 10) {
								break;
							}
						}
					}
					break;
				}
			}
			csvPrinter.close();

		} catch (ColsNotFoundException e) {
			response.setStatus(400);
			response.getWriter().println(e.getMessage());
			logger.debug(e.getMessage());
			e.printStackTrace();
		} catch (Exception e) {
			response.setStatus(400);
			logger.debug("ODS to CSV fail. Exception: " + e.toString());
			e.printStackTrace();
		}

		return;
	}

	// TODO: not recognizing currency???
	// helper for ODStoCSV
	private static String getOdsCellVal(org.jopendocument.dom.spreadsheet.Cell<SpreadSheet> cell) {
		if (!cell.isEmpty()) {
			switch (cell.getValueType()) {
			case CURRENCY:
				NumberFormat formatter = NumberFormat.getCurrencyInstance();
				String curr = formatter.format(Double.valueOf(cell.getValue().toString())).trim();
				logger.debug(curr);
				return curr;
			// return cell.getValue().toString();
			// return Double.parseDouble(cell.getTextValue()) + "";
			case STRING:
				return cell.getValue() + "";
			case DATE:
				SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
				return sdf.format(cell.getValue());
			// return cell.getValue().toString();
			case FLOAT:
				// logger.debug(cell.getValue() + "");
				return cell.getValue() + "";
			default:
				return cell.getValue() + "";
			}
		}
		return "";
	}

	private static void XMLtoXINI(File infile, String outfile, HttpServletResponse response) {
		// convert the XML document to custom XPath INI
		XMLToXPathIniSerializer xPathIniSerializer = new XMLToXPathIniSerializer();
		xPathIniSerializer.setFeature(XMLToXPathIniSerializer.FEATURE_WRITE_ATTRIBUTES, true);
		xPathIniSerializer.setFeature(XMLToXPathIniSerializer.FEATURE_WRITE_ROOT_ELEMENT, true);
		try {
			DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
			Document doc = docBuilder.parse(infile);

			PrintStream fileOut = new PrintStream(outfile);
			xPathIniSerializer.writeToStream(doc, "", fileOut);
		} catch (Exception e) {
			response.setStatus(400);
			logger.debug("XML to XINI fail. Exception: " + e.toString());
			e.printStackTrace();
		}

		return;
	}

	private static File tempCSV(String filename, HttpServletResponse response) {
		File tempFile;
		try {
			tempFile = new File(filename);
			tempFile.deleteOnExit();
			logger.debug("Temp csv file is: " + tempFile.getAbsolutePath());
			return tempFile;
		} catch (Exception e) {
			response.setStatus(400);
			logger.debug("Temp csv creation fail. Exception: " + e.toString());
			e.printStackTrace();
		}
		return null;
	}
}
