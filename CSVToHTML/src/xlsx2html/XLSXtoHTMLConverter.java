/**************************************************************************
 * @author Avinash Vablisetty(vnvak1994@gmail.com)
 * 
 * This application is used to convert an .XLSX file to .HTML file
 * The content in the XLSX files remains same in the generated HTML file and options to navigate between sheets will also be available
 * Apache POI Jar has been used to convert the files 
 * 
 * 
 ***********************************************************************/

package xlsx2html;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFSheet;

import java.io.*;

public class XLSXtoHTMLConverter {
	public static void main(String[] args) {
		String projectRoot = System.getProperty("user.dir");

		String xlsFilePath = "\\XLSXFiles\\SampleXLSFile_19kb.xls";
		String htmlFilePath = "\\OutputFile\\output.html";

		String xlsFile = projectRoot + xlsFilePath;
		String htmlFile = projectRoot + htmlFilePath;

		// Use the absolute paths in your code
		System.out.println("XLS File Path: " + xlsFile);
		System.out.println("HTML File Path: " + htmlFile);

		try {
			FileInputStream fis = new FileInputStream(xlsFile);
			Workbook workbook = new HSSFWorkbook(fis);

			FileOutputStream fos = new FileOutputStream(htmlFile);
			PrintWriter pw = new PrintWriter(fos);

			pw.println("<html>\n<body>");

			// Generate the container div to display the sheets
			pw.println("<div id=\"sheetsContainer\"></div>");

			// JavaScript for tab functionality
			pw.println("<script>");
			pw.println("var sheetsContainer = document.getElementById('sheetsContainer');");
			pw.println("function showSheet(sheetName) {");
			pw.println("  var sheets = document.getElementsByClassName('sheet');");
			pw.println("  for (var i = 0; i < sheets.length; i++) {");
			pw.println("    sheets[i].style.display = 'none';");
			pw.println("  }");
			pw.println("  var sheet = document.getElementById(sheetName);");
			pw.println("  sheet.style.display = 'block';");
			pw.println("}");
			pw.println("document.addEventListener('DOMContentLoaded', function() {");
			pw.println("  var sheets = document.getElementsByClassName('sheet');");
			pw.println("  for (var i = 1; i < sheets.length; i++) {");
			pw.println("    sheets[i].style.display = 'none';");
			pw.println("  }");
			pw.println("});");
			pw.println("</script>");

			// Generate the content for each sheet
			for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
				HSSFSheet sheet = (HSSFSheet) workbook.getSheetAt(i);
				String sheetName = sheet.getSheetName();

				pw.println("<div id=\"" + sheetName + "\" class=\"sheet\">");
				pw.println("<h2>Sheet: " + sheetName + "</h2>");
				pw.println("<table border='1'>");

				for (Row row : sheet) {
					pw.println("<tr>");

					for (Cell cell : row) {
						pw.println("<td>");

						switch (cell.getCellType()) {
						case STRING:
							pw.println(cell.getRichStringCellValue().getString());
							break;
						case NUMERIC:
							if (DateUtil.isCellDateFormatted(cell)) {
								pw.println(cell.getDateCellValue());
							} else {
								pw.println(cell.getNumericCellValue());
							}
							break;
						case BOOLEAN:
							pw.println(cell.getBooleanCellValue());
							break;
						case FORMULA:
							pw.println(cell.getCellFormula());
							break;
						default:
							pw.println("");
						}

						pw.println("</td>");
					}

					pw.println("</tr>");
				}

				pw.println("</table>");
				pw.println("</div>");
			}

			// Generate the tab buttons at the bottom
			pw.println("<div style=\"text-align: left;\">");
			for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
				HSSFSheet sheet = (HSSFSheet) workbook.getSheetAt(i);
				String sheetName = sheet.getSheetName();

				pw.println("<button onclick=\"showSheet('" + sheetName + "')\">" + sheetName + "</button>");
			}
			pw.println("</div>");

			pw.println("</body>\n</html>");

			pw.close();
			fos.close();

			System.out.println("XLS to HTML conversion complete!");

			fis.close();

		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
