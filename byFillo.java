package excelCompare;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.codoid.products.fillo.Connection;
import com.codoid.products.fillo.Fillo;
import com.codoid.products.fillo.Recordset;

public class byFillo 
{
	static XSSFSheet VICCI_sheet;
	static XSSFSheet VICTOR_sheet;
	public static void main(String[] args) throws Exception 
	{	
		//  Excel workbook initialization using APACHE POI
		
		File vicciData = new File("C:\\Users\\Sankalp\\Desktop\\VICCI.xlsx");
		File victorData = new File("C:\\Users\\Sankalp\\Desktop\\VICTOR.xlsx");
		FileInputStream VICCI_File = new FileInputStream(vicciData);
		FileInputStream VICTOR_File = new FileInputStream(victorData);
		XSSFWorkbook VICCIBook = new XSSFWorkbook(VICCI_File);
		XSSFWorkbook VICTORBook = new XSSFWorkbook(VICTOR_File);
		VICCI_sheet = VICCIBook.getSheetAt(0);
		VICTOR_sheet = VICTORBook.getSheetAt(0);
		
		// Values for iteration
		
		int vicciCarlines = getRows(VICCI_sheet);
		int vicciColumns = getCell(0, VICCI_sheet);
		
		// Variables used
		
		Object CarLineData_VICTOR = null;
		Object SalesGroup_VICTOR = null;
		Object Model_VICTOR = null;
		Object Missing_VICTOR = null;
		Connection compareConnection;
		String query = null;
		boolean victorCarLineMissing = false;
		String fileName;
		
		// Arrays used
		
		HashMap<Integer, Integer> vicciKeys = new HashMap<Integer, Integer>();
		List<Integer> vicciFinalKey = new ArrayList<Integer>();
		
		// Fillo class initialization
		
		Fillo fillo = new Fillo();
		Connection connectionVICCI = fillo.getConnection("C:\\Users\\Sankalp\\Desktop\\VICCI.xlsx");
		Connection connectionVICTOR = fillo.getConnection("C:\\Users\\Sankalp\\Desktop\\VICTOR.xlsx");
		Recordset recordsetVICCI;
		Recordset recordsetVICTOR = null;
		
		// New workbook creation for storing comparison data
		
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet firstSheet = wb.createSheet("Sheet1");
		XSSFRow row = firstSheet.createRow(0);
		XSSFCell cellOne = row.createCell(0);
		cellOne.setCellValue(new XSSFRichTextString("CarLine"));
		cellOne.setCellType(CellType.STRING);
		XSSFCell cellTwo = row.createCell(1);
		cellTwo.setCellValue(new XSSFRichTextString("SalesGroup"));
		XSSFCell cellThree = row.createCell(2);
		cellThree.setCellValue(new XSSFRichTextString("Model"));
		XSSFCell cellFour = row.createCell(3);
		cellFour.setCellValue(new XSSFRichTextString("MissingFromVictor"));
		
		//Creating excel file with label as current time stamp
		
		fileName=getFileName();
		try (FileOutputStream fos = new FileOutputStream(new File("C:\\Users\\Sankalp\\Desktop\\" + fileName + ".xlsx"))) {
			wb.write(fos);
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		// Establishing connection to comparison excel file
		
		compareConnection = fillo.getConnection("C:\\Users\\Sankalp\\Desktop\\" + fileName + ".xlsx");
		
		// Fetching all rows for 0th cell from VICCI file & storing in hashmap 
		
		for (int i = 0; i < 1; i++) {
			for (int j = 1; j <= vicciCarlines; j++) 
			{
				String no = VICCI_sheet.getRow(j).getCell(i).toString();

				int data = (int) VICCI_sheet.getRow(j).getCell(i).getNumericCellValue();

				// int keys = Integer.parseInt(no);

				if (vicciKeys.containsKey(data)) {

				} else {
					vicciKeys.put(data, data);
				}
			}
		}
		
		// Iterating hashmap to get all keys from it


		for (Integer chabi : vicciKeys.keySet()) 
		{
			vicciFinalKey.add(chabi);
		}
		
		// Iterating over all the unique keys from file 1
		
		for (int i = 0; i < vicciFinalKey.size(); i++) {
			try {
				//System.out.println(vicciFinalKey.get(i));
				
				// Saving all records for each key using select query for both VICCI & VICTOR data

				recordsetVICTOR = connectionVICTOR.executeQuery("Select * from Sheet1")
						.where("CarLine='" + vicciFinalKey.get(i) + "'");
				recordsetVICCI = connectionVICCI.executeQuery("Select * from Sheet1")
						.where("CarLine='" + vicciFinalKey.get(i) + "'");
				
				// Iterating over each record for current key one by one
				
				while (recordsetVICTOR.next()) {
					
					// If car line is missing then in catch block all fields will be written as blank in result excel

					CarLineData_VICTOR = recordsetVICTOR.getField("CarLine");
					
					//Checking if sales group is available in VICTOR data for current car line key

					if (recordsetVICTOR.getField("SalesGroup").isBlank()) 
					{

						// If blank then logically no data will be present for model so putting values of sales group& model as blank in excel
						
						SalesGroup_VICTOR = "Blank";
						Model_VICTOR = "Blank";

						// Taking all records for first key from VICCI to write write missing sales group from VICTOR in result file
						
						recordsetVICCI = connectionVICCI.executeQuery("Select * from Sheet1")
								.where("CarLine='" + vicciFinalKey.get(i) + "'");

						while (recordsetVICCI.next()) {
							Missing_VICTOR = recordsetVICCI.getField("SalesGroup");
							System.out.println("SalesGroup missing from VICTOR is" + Missing_VICTOR);
						}

						query = "INSERT INTO Sheet1(CarLine,SalesGroup,Model,MissingFromVictor) VALUES('"
								+ CarLineData_VICTOR + "'," + "'" + SalesGroup_VICTOR + "'," + "'" + Model_VICTOR + "',"
								+ "'" + Missing_VICTOR + "')";
						compareConnection.executeUpdate(query);

					} else 
					{
						// If models are not available in VICTOR then marking it as blank in result excel
						if (recordsetVICTOR.getField("Model").isBlank()) {
							// SalesGroup_VICTOR = "Blank";
							Model_VICTOR = "Blank";
							recordsetVICCI = connectionVICCI.executeQuery("Select * from Sheet1")
									.where("CarLine='" + vicciFinalKey.get(i) + "'");
							while (recordsetVICCI.next()) {
								Missing_VICTOR = recordsetVICCI.getField("Model");
								//System.out.println("Model missing from VICTOR is:" + Missing_VICTOR);
							}

							query = "INSERT INTO Sheet1(CarLine,SalesGroup,Model,MissingFromVictor) VALUES('"
									+ CarLineData_VICTOR + "'," + "'" + SalesGroup_VICTOR + "'," + "'" + Model_VICTOR
									+ "'," + "'" + Missing_VICTOR + "')";
							compareConnection.executeUpdate(query);

						} else {
							SalesGroup_VICTOR = recordsetVICTOR.getField("SalesGroup");
							// System.out.println(SalesGroup_VICTOR);
							Model_VICTOR = recordsetVICTOR.getField("Model");

						}

						//System.out.println(CarLineData_VICTOR + "...." + SalesGroup_VICTOR + "...." + Model_VICTOR);
						//System.out.println("inserting to excel");

						
					}

				}
			} catch (Exception E) 
			{
				// If Car line (Key) from VICCI is not present in VICTOR then writing all cells as NA in result excel along with missing key
				
				victorCarLineMissing = true;

				if (victorCarLineMissing == true) {
					compareConnection = fillo.getConnection("C:\\Users\\Sankalp\\Desktop\\" + fileName + ".xlsx");

					recordsetVICCI = connectionVICCI.executeQuery("Select * from Sheet1")
							.where("CarLine='" + vicciFinalKey.get(i) + "'");

					while (recordsetVICCI.next()) {
						Missing_VICTOR = recordsetVICCI.getField("CarLine");
						//System.out.println(Missing_VICTOR);
					}
					
					// Insert query to insert values for particular field in result excel

					query = "INSERT INTO Sheet1(CarLine,SalesGroup,Model,MissingFromVictor) VALUES('NA','NA','NA',"
							+ Missing_VICTOR + ")";

					compareConnection.executeUpdate(query);

					//System.out.println("Writing to excel");

					//VICTOR_Finals.add(new Object[] { "No data", "No data", "No data" });

					victorCarLineMissing = false;
				}

			}

		}

		compareConnection.close();
		wb.close();
		VICCIBook.close();
		VICTORBook.close();
	}
	
	// Method to fetch all rows from provided sheet
	
	public static int getRows(XSSFSheet sheet) {
		int rowCount = sheet.getLastRowNum();

		return rowCount;
	}
	
	// Method to fetch all cells from provided sheet & row

	public static int getCell(int row, XSSFSheet sheet) {
		int cells = sheet.getRow(row).getLastCellNum();

		return cells;
	}

	// Method to generate file name
	
	public static String getFileName()
	{
		Date dt = new Date();

		String dt1 = dt.toString();

		String dt2 = dt1.replaceAll("\\s", "_");

		String dt3 = dt2.replaceAll(":", "_");

		System.out.println(dt3);
		
		return dt3;

	}

}
