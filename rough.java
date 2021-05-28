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
		Recordset  recordsetVICCI = null;
		Recordset recordsetVICTOR = null;
		
		
		List<Object[]> VICTOR_Finals = new ArrayList<Object[]>();
		List<Object> arrlist1_VICTOR = null;
		Object CarLineData_VICTOR;
		Object SalesGroup_VICTOR;
		Object Model_VICTOR;
		
		for (int i = 0; i < vicciFinalKey.size(); i++)
		{
			try
			{
				System.out.println(vicciFinalKey.get(i));
				
				recordsetVICTOR = connectionVICTOR.executeQuery("Select * from Sheet1").where("CarLine='" + vicciFinalKey.get(i) + "'");
				while(recordsetVICTOR.next())
				{
					if (recordsetVICTOR.getField("CarLine").isBlank()) {
						CarLineData_VICTOR = "Blank";
					} else {
						CarLineData_VICTOR = recordsetVICTOR.getField("CarLine");
					}

					// ******************

					if (recordsetVICTOR.getField("SalesGroup").isBlank()) {
						SalesGroup_VICTOR = "Blank";
					} else {
						SalesGroup_VICTOR = recordsetVICTOR.getField("SalesGroup");
					}

					// *******************

					if (recordsetVICTOR.getField("Model").isBlank()) {
						Model_VICTOR = "Blank";
					} else {
						Model_VICTOR = recordsetVICTOR.getField("Model");
					}
					
					VICTOR_Finals.add(new Object[] { CarLineData_VICTOR, SalesGroup_VICTOR, Model_VICTOR });
				}
			}
			catch(Exception E)
			{
				VICTOR_Finals.add(new Object[] {"For the Car Line from VICCI"+" "+vicciFinalKey.get(i)+" "+"No records available in VICTOR"});
				
			}
		}
		
		for (int i = 0; i < VICTOR_Finals.size(); i++) {

			arrlist1_VICTOR = new ArrayList<>(Arrays.asList(VICTOR_Finals.get(i)));

			System.out.println(arrlist1_VICTOR);
		}
		
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

			recordsetVICCI = connectionVICCI.executeQuery("Select * from Sheet1").where("CarLine='" + vicciFinalKey.get(i) + "'");

			//System.out.println("Total records retrived are:" + " " + recordsetVICCI.getCount());

			if (recordsetVICCI.getCount() != 0) {
				while (recordsetVICCI.next()) {
					// int index=1;

					/*System.out.println(recordsetVICCI.getField("CarLine") + "...." + recordsetVICCI.getField("SalesGroup")
							+ "...." + recordsetVICCI.getField("Model"));*/

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
				
				VICCIFinals.add(new Object[] {"For the Car Line"+" "+vicciFinalKey.get(i)+" "+"No records available in VICTOR"});

			}

		}

		for (int i = 0; i < VICCIFinals.size(); i++) {

			arrlist1_VICCI = new ArrayList<>(Arrays.asList(VICCIFinals.get(i)));

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
