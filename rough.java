import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.codoid.products.fillo.Connection;
import com.codoid.products.fillo.Fillo;
import com.codoid.products.fillo.Recordset;

public class rough {
	static XSSFSheet file1_sheet;
	static XSSFSheet file2_sheet;

	public static void main(String[] args) throws Exception {
		// Fetching file
		File file1Data = new File("C:\\Users\\Sankalp\\Desktop\\VICCI.xlsx");
		File file2Data = new File("C:\\Users\\Sankalp\\Desktop\\VICTOR.xlsx");

		FileInputStream file1 = new FileInputStream(file1Data);
		FileInputStream file2 = new FileInputStream(file2Data);
		
		//Initiliazing workbook

		XSSFWorkbook file1Book = new XSSFWorkbook(file1);
		XSSFWorkbook file2Book = new XSSFWorkbook(file2);
		
		// Pointing to required sheet index

		file1_sheet = file1Book.getSheetAt(0);
		file2_sheet = file2Book.getSheetAt(0);
		
		// Fetching filled rows & cells from file1

		int file1_RowCount = getRows(file1_sheet);
		int primaryRowCells = getCell(0, file1_sheet);
		
		// Till now Apache POI libraries has been used due to requirement

		HashMap<Integer, Integer> file1Keys = new HashMap<Integer, Integer>();
		
		/********** Processing for file2   ************/ 
		
		// Storing unique keys for all rows && first cell

		for (int i = 0; i < 1; i++) {
			for (int j = 1; j <= file1_RowCount; j++) {
				String no = file1_sheet.getRow(j).getCell(i).toString();

				int data = (int) file1_sheet.getRow(j).getCell(i).getNumericCellValue();

				// int keys = Integer.parseInt(no);

				if (file1Keys.containsKey(data)) {

				} else {
					file1Keys.put(data, data);
				}
			}
		}

		List<Integer> file1FinalKeys = new ArrayList<Integer>();

		//Adding filtered unique keys of first file to list

		for (Integer chabi : file1Keys.keySet()) {
			// int i=0;

			file1FinalKeys.add(chabi);

			// i++;
		}
		
		// Initilizing Fillo class and passing file path

		Fillo fillo = new Fillo();
		Connection connectionFile1 = fillo.getConnection("C:\\Users\\Sankalp\\Desktop\\file1.xlsx");
		Connection connectionFile2 = fillo.getConnection("C:\\Users\\Sankalp\\Desktop\\file2.xlsx");
		
		//Recordset will store result of passing query
		Recordset  recordsetFile1 = null;
		Recordset recordsetFile2 = null;
		
		
		List<Object[]> file2_Finals = new ArrayList<Object[]>();
		List<Object> arrlist1_file2 = null;
		Object exactColumnName1_file2;
		Object exactColumnName2_file2;
		Object exactColumnName3_file2;
		
		for (int i = 0; i < file1FinalKeys.size(); i++)
		{
			try
			{
				
				
				recordsetFile2 = connectionVICTOR.executeQuery("Select * from Sheet1").where("Required_column_name_from_Excel2='" + file1FinalKeys.get(i) + "'");
				while(recordsetFile2.next())
				{
					if (recordsetFile2.getField("column1Name").isBlank()) {
						exactColumnName1_file2 = "Blank";
					} else {
						exactColumnName1_file2 = recordsetFile2.getField("column1Name");
					}

					// ******************

					if (recordsetFile2.getField("column2Name").isBlank()) {
						exactColumnName2_file2 = "Blank";
					} else {
						exactColumnName2_file2 = recordsetFile2.getField("column2Name");
					}

					// *******************

					if (recordsetFile2.getField("column3Name").isBlank()) {
						exactColumnName3_file2 = "Blank";
					} else {
						exactColumnName3_file2 = recordsetFile2.getField("column3Name");
					}
					
					file2_Finals.add(new Object[] { column1Name, column2Name, column3Name });
				}
			}
			catch(Exception E)
			{
				file2_Finals.add(new Object[] {"For the key from file1"+" "+file1FinalKeys.get(i)+" "+"No records available in file2"});
				
			}
		}
		
		for (int i = 0; i < file2_Finals.size(); i++) {

			arrlist1_file2 = new ArrayList<>(Arrays.asList(file2_Finals.get(i)));

			System.out.println(arrlist1_file2);
		}
		
		
		
		List<Object[]> Finals = new ArrayList<Object[]>();
		List<Object> arrlist1 = null;
		Object fiile1cell1;
		Object fiile1cell2;
		Object fiile1cell3;

		for (int i = 0; i < file1FinalKeys.size(); i++) {
			String strQuery = "Select * from Sheet1 where CarLine where CarLine='" + file1FinalKeys.get(i) + "'";
			// System.out.println("First key is:"+vicciFinalKey.get(i));

			// System.out.println("First input query is:");
			// System.out.println();
			// System.out.println(strQuery);

			recordsetFile1 = connection.executeQuery("Select * from Sheet1").where("CarLine='" + file1FinalKeys.get(i) + "'");

			//System.out.println("Total records retrived are:" + " " + recordsetVICCI.getCount());

			if (recordsetFile1.getCount() != 0) {
				while (recordsetFile2.next()) {
					// int index=1;

					

					if (recordsetFile1.getField("columnName").isBlank()) {
						fiile1cell1 = "Blank";
					} else {
						fiile1cell1 = recordsetFile1.getField("columnName");
					}

					// ******************

					if (recordsetFile1.getField("columnName").isBlank()) {
						fiile1cell2 = "Blank";
					} else {
						fiile1cell2 = recordsetFile1.getField("columnName");
					}

					// *******************

					if (recordset.getField("columnName").isBlank()) {
						fiile1cell3 = "Blank";
					} else {
						fiile1cell3 = recordsetFile1.getField("columnName");
					}

					Finals.add(new Object[] { fiile1cell1, fiile1cell2, fiile1cell3 });
				}
			} else {
				
				Finals.add(new Object[] {"For the key "+" "+FinalKey.get(i)+" "+"No records available in file2"});

			}

		}

		for (int i = 0; i < Finals.size(); i++) {

			arrlist1_file2 = new ArrayList<>(Arrays.asList(Finals.get(i)));

			System.out.println(arrlist1_VICCI);
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
