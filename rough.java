import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.codoid.products.exception.FilloException;
import com.codoid.products.fillo.Connection;
import com.codoid.products.fillo.Fillo;
import com.codoid.products.fillo.Recordset;

public class rough {
	static XSSFSheet VICCI_sheet;
	static XSSFSheet VICTOR_sheet;

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub

		File vicciData = new File("C:\\Users\\Sankalp\\Desktop\\VICCI.xlsx");
		File victorData = new File("C:\\Users\\Sankalp\\Desktop\\VICTOR.xlsx");

		FileInputStream VICCI_File = new FileInputStream(vicciData);
		FileInputStream VICTOR_File = new FileInputStream(victorData);

		XSSFWorkbook VICCIBook = new XSSFWorkbook(VICCI_File);
		XSSFWorkbook VICTORBook = new XSSFWorkbook(VICTOR_File);

		VICCI_sheet = VICCIBook.getSheetAt(0);
		VICTOR_sheet = VICTORBook.getSheetAt(0);

		int vicciCarlines = getRows(VICCI_sheet);
		int vicciColumns = getCell(0, VICCI_sheet);

		HashMap<Integer, Integer> vicciKeys = new HashMap<Integer, Integer>();

		for (int i = 0; i < 1; i++) {
			for (int j = 1; j <= vicciCarlines; j++) {
				String no = VICCI_sheet.getRow(j).getCell(i).toString();

				int data = (int) VICCI_sheet.getRow(j).getCell(i).getNumericCellValue();

				// int keys = Integer.parseInt(no);

				if (vicciKeys.containsKey(data)) {

				} else {
					vicciKeys.put(data, data);
				}
			}
		}

		List<Integer> vicciFinalKey = new ArrayList<Integer>();

		// System.out.println(vicciKeys);

		for (Integer chabi : vicciKeys.keySet()) {
			// int i=0;

			vicciFinalKey.add(chabi);

			// i++;
		}

		Fillo fillo = new Fillo();
		Connection connectionVICCI = fillo.getConnection("C:\\Users\\Sankalp\\Desktop\\VICCI.xlsx");
		Connection connectionVICTOR = fillo.getConnection("C:\\Users\\Sankalp\\Desktop\\VICTOR.xlsx");
		Recordset recordsetVICCI = null;
		Recordset recordsetVICTOR = null;
		boolean vicciCarLineMissing = false;
		boolean vicciSalesGroupMissing = false;
		boolean vicciModel = false;

		List<Object[]> VICTOR_Finals = new ArrayList<Object[]>();
		List<Object> arrlist1_VICTOR = null;
		Object CarLineData_VICTOR = null;
		Object SalesGroup_VICTOR = null;
		Object Model_VICTOR = null;
		boolean victorCarLineMissing = false;
		boolean victorSalesGroupMissing = false;
		boolean victorModel = false;
		Connection compareConnection;
		String query = null;

		// creating an instance of Workbook class
		XSSFWorkbook wb = new XSSFWorkbook();
		// creates an excel file at the specified location

		Date dt = new Date();

		String dt1 = dt.toString();

		String dt2 = dt1.replaceAll("\\s", "_");

		String dt3 = dt2.replaceAll(":", "_");

		System.out.println(dt3);

		XSSFSheet firstSheet = wb.createSheet("Sheet1");

		XSSFRow row = firstSheet.createRow(0);
		XSSFCell cellOne = row.createCell(0);
		cellOne.setCellValue(new XSSFRichTextString("CarLine"));

		XSSFCell cellTwo = row.createCell(1);
		cellTwo.setCellValue(new XSSFRichTextString("SalesGroup"));

		XSSFCell cellThree = row.createCell(2);
		cellThree.setCellValue(new XSSFRichTextString("Model"));

		try (FileOutputStream fos = new FileOutputStream(new File("C:\\Users\\Sankalp\\Desktop\\" + dt3 + ".xlsx"))) {
			wb.write(fos);
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		compareConnection = fillo.getConnection("C:\\Users\\Sankalp\\Desktop\\" + dt3 + ".xlsx");

		for (int i = 0; i < vicciFinalKey.size(); i++) {
			try {
				System.out.println(vicciFinalKey.get(i));

				recordsetVICTOR = connectionVICTOR.executeQuery("Select * from Sheet1")
						.where("CarLine='" + vicciFinalKey.get(i) + "'");
				recordsetVICCI = connectionVICCI.executeQuery("Select * from Sheet1")
						.where("CarLine='" + vicciFinalKey.get(i) + "'");
				while (recordsetVICTOR.next()) {

					CarLineData_VICTOR = recordsetVICTOR.getField("CarLine");

					// ******************
					
					System.out.println("Current salesgroup:"+" "+recordsetVICTOR.getField("SalesGroup"));

					if (recordsetVICTOR.getField("SalesGroup").isBlank()
							) 
					{
						SalesGroup_VICTOR = "Blank";
						Model_VICTOR = "Blank";

						// SalesGroup_VICTOR = "Blank";
						victorSalesGroupMissing = true;
						victorModel = true;
						vicciSalesGroupMissing = true;
						vicciModel = true;

					} else 
					{
						//SalesGroup_VICTOR = recordsetVICTOR.getField("SalesGroup");
						//Model_VICTOR = recordsetVICTOR.getField("Model");

						if (recordsetVICTOR.getField("Model").isBlank() ) 
						{
							//SalesGroup_VICTOR = "Blank";
							Model_VICTOR = "Blank";

						} else 
						{
							SalesGroup_VICTOR = recordsetVICTOR.getField("SalesGroup");
							//System.out.println(SalesGroup_VICTOR);
							Model_VICTOR = recordsetVICTOR.getField("Model");
							
						}
						
						System.out.println(CarLineData_VICTOR+"...."+SalesGroup_VICTOR+"...."+Model_VICTOR);
						System.out.println("inserting to excel");
						
						victorSalesGroupMissing = false;
						victorModel = false;
						vicciSalesGroupMissing = false;
						vicciModel = false;
					}

					

					 query = "INSERT INTO Sheet1(CarLine,SalesGroup,Model) VALUES('" + CarLineData_VICTOR + "',"
							+ "'" + SalesGroup_VICTOR + "'," + "'" + Model_VICTOR + "')";


					VICTOR_Finals.add(new Object[] { CarLineData_VICTOR, SalesGroup_VICTOR, Model_VICTOR });

				}
			} catch (Exception E) {
				victorCarLineMissing = true;

				if (victorCarLineMissing == true) 
				{
					compareConnection = fillo.getConnection("C:\\Users\\Sankalp\\Desktop\\" + dt3 + ".xlsx");
					
					 query = "INSERT INTO Sheet1(CarLine,SalesGroup,Model) VALUES('Blank','Blank','Blank')";

					System.out.println("Writing to excel");

					VICTOR_Finals.add(new Object[] { "No data", "No data", "No data" });
					
					victorCarLineMissing=false;
				}

			}
			
		}
		compareConnection.executeUpdate(query);

		compareConnection.close();
		
		for (int i = 0; i < VICTOR_Finals.size(); i++) {

			arrlist1_VICTOR = new ArrayList<>(Arrays.asList(VICTOR_Finals.get(i)));

			System.out.println(arrlist1_VICTOR);

			
		}

		// OutputStream vicci_VICTOR_Compare = new
		// FileOutputStream("C:\\Users\\Sankalp\\Desktop\\dt3.xlsx");

		// System.out.println("Excel File has been created successfully.");
		// wb.write(fileOut);

		List<Object[]> VICCIFinals = new ArrayList<Object[]>();
		List<Object> arrlist1_VICCI = null;
		Object CarLineData_VICCI;
		Object SalesGroup_VICCI;
		Object Model_VICCI;

		for (int i = 0; i < vicciFinalKey.size(); i++) {
			String strQuery = "Select * from Sheet1 where CarLine where CarLine='" + vicciFinalKey.get(i) + "'";
			// System.out.println("First key is:"+vicciFinalKey.get(i));

			// System.out.println("First input query is:");
			// System.out.println();
			// System.out.println(strQuery);

			recordsetVICCI = connectionVICCI.executeQuery("Select * from Sheet1")
					.where("CarLine='" + vicciFinalKey.get(i) + "'");

			// System.out.println("Total records retrived are:" + " " +
			// recordsetVICCI.getCount());

			if (recordsetVICCI.getCount() != 0) {
				while (recordsetVICCI.next()) {
					// int index=1;

					/*
					 * System.out.println(recordsetVICCI.getField("CarLine") + "...." +
					 * recordsetVICCI.getField("SalesGroup") + "...." +
					 * recordsetVICCI.getField("Model"));
					 */

					if (recordsetVICCI.getField("CarLine").isBlank()) {
						CarLineData_VICCI = "Blank";
					} else {
						CarLineData_VICCI = recordsetVICCI.getField("CarLine");
					}

					// ******************

					if (recordsetVICCI.getField("SalesGroup").isBlank()) {
						SalesGroup_VICCI = "Blank";
					} else {
						SalesGroup_VICCI = recordsetVICCI.getField("SalesGroup");
					}

					// *******************

					if (recordsetVICCI.getField("Model").isBlank()) {
						Model_VICCI = "Blank";
					} else {
						Model_VICCI = recordsetVICCI.getField("Model");
					}

					VICCIFinals.add(new Object[] { CarLineData_VICCI, SalesGroup_VICCI, Model_VICCI });
				}
			} else {

				VICCIFinals.add(new Object[] {
						"For the Car Line" + " " + vicciFinalKey.get(i) + " " + "No records available in VICTOR" });

			}

		}

		for (int i = 0; i < VICCIFinals.size(); i++) {

			arrlist1_VICCI = new ArrayList<>(Arrays.asList(VICCIFinals.get(i)));

			System.out.println(arrlist1_VICCI);

			System.out.println("Next array");
		}

	}

	public static int getRows(XSSFSheet sheet) {
		int rowCount = sheet.getLastRowNum();

		return rowCount;
	}

	public static int getCell(int row, XSSFSheet sheet) {
		int cells = sheet.getRow(row).getLastCellNum();

		return cells;
	}

}
