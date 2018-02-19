package wcec.excel_parser;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Hashtable;
import java.util.List;

import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * On Sheet 0 of the workbook - 
 * 
 * Cell[6, 0] = Joy	    Cell[6, 1] = Gentle	    Cell[6, 2] = Peace	     Cell[6, 3] = Cell[6, 4] = Patience	Cell[6, 5] = Abundance	Cell[6, 6] = Self-Control	Cell[6, 7] = Love	    Cell[6, 8] = Kindness	Cell[6, 9] = Faith	    Cell[6, 10] = Cantonese	Cell[6, 11] = Goodness
 * Cell[7, 0] = 喜樂小組	Cell[7, 1] = 溫柔小組	Cell[7, 2] = 和平小組	Cell[7, 3] = Cell[7, 4] = 忍耐小組	Cell[7, 5] = 豐盛小組	Cell[7, 6] = 節制小組	    Cell[7, 7] = 仁愛小組	Cell[7, 8] = 恩慈小組	Cell[7, 9] = 信實小組	Cell[7, 10] = 粵語團契	Cell[7, 11] = 良善小組	 
 * Cell[9, 0] = 石　翔 陳浣方	Cell[9, 1] = 金一鸣 胡　敏	Cell[9, 2] = 方健雄 吳綏歐	Cell[9, 3] = Cell[9, 4] = 于　斌 劉　穎	Cell[9, 5] = 王深義  萬湘英	Cell[9, 6] = 莊志豪  李　展	Cell[9, 7] = 陳源石 鄭炎嫻	Cell[9, 8] = 張仁榮  陳岱純	Cell[9, 9] = 沈自強 林素芳	Cell[9, 10] = 朱　海 鄭穎瑜	Cell[9, 11] = 赵　航   潘智迎 
 * Cell[10, 0] = 494-8150	Cell[10, 1] = 610-558-1318	Cell[10, 2] =  239-7701	Cell[10, 3] = Cell[10, 4] = 729-2704	Cell[10, 5] = 369-0694	Cell[10, 6] = 294-0104	Cell[10, 7] = 266-7981	Cell[10, 8] = 766-1602	Cell[10, 9] = 234-8406	Cell[10, 10] = 722-4719	Cell[10, 11] = 301-202-4845	 
 * Cell[55, 0] = Cell[55, 1] = Cell[55, 5] = Cell[55, 6] = Cell[55, 8] = Cell[55, 9] = Cell[55, 10] = Cell[55, 11] = Cell[55, 12] = Cell[55, 13] = Cell[55, 17] = 

 */
public class App 
{
	static int CellGroupRow = 7;           // row 7 contains the name of cell group
	static int NumberOfCellGroups = 12;    // 11 groups
	
	static int LeaderRow = 9;
	static int MemberStartingRow = 12;
	
	static int TotalGroupCount = 0;
	
	static Integer JoyGroupColumn = 0;
	static Integer GentleGroupColumn = 1;
	static Integer PeaceGroupColumn = 2;
	static Integer PeaceGroupColumn2 = 3;
	static Integer PatienceGroupColumn = 4;
	static Integer AbundanceGroupColumn = 5;
	static Integer SelControlGroupColumn = 6;
	static Integer LoveGroupColumn = 7;
	static Integer KindnessGroupColumn = 8;
	static Integer FaithGroupColumn =  9;
	static Integer CantoneseGroupColumn = 10;
	static Integer GoodnessGroupColumn = 11;
	
	static List<Integer> TheGroupColumns = new ArrayList<Integer>();
	
	static final int Undefined = -1;
	static final int SingleMemberWithPhone = 1;
	static final int SingleMemberWithoutPhone = 2;
	static final int BothMemberWithPhone = 3;
	
	public static String TheFileName = "C:/tmp/church_directory.xls";

	public static void readFromExcelFile(String[] args) {
		Hashtable<Integer, CellGroup> allGroups = new Hashtable<Integer, CellGroup>();
		
	    try {
	    	File aFile = new File (TheFileName);
	    	// SS Workbook object
	    	Workbook workbook = WorkbookFactory.create(new FileInputStream(aFile)); 
	    	//Get first/desired sheet from the workbook
	        Sheet sheet = workbook.getSheetAt(0);
	        // allGroups will have all the cell groups
	        populateCellGroups(sheet, allGroups); 
	        for (int i = 0; i < TheGroupColumns.size(); i++) {
	        	CellGroup currentGroup = allGroups.get(i);
	        	Row leaderRow = sheet.getRow(LeaderRow);
	        	Cell leaderCell = leaderRow.getCell(i);
	        	Family newFamily = new Family();
	        	parseCell(sheet, leaderCell, newFamily);
	        	currentGroup.setLeadFamily(newFamily);
	        	// continue to read till end of row
	        	for (int r = MemberStartingRow; r < 100; r++) {
	        		Row aRow = sheet.getRow(r);
		        	Cell aCell = aRow.getCell(i);
		        	Family memberFamily = new Family();
		        	int result = parseCell(sheet, aCell, memberFamily);
		        	switch (result) {
		        	case BothMemberWithPhone:
		        		currentGroup.addMemberFamily(memberFamily); 
		        		r++;
		        		break; 
		        	case SingleMemberWithPhone:
		        		currentGroup.addMemberFamily(memberFamily); 
		        		break;
		        	case SingleMemberWithoutPhone: // sometimes the phone number is on the next line for this case - TBD
		        		break;
		        	case Undefined:
		        		int j = 1;
		        		j++;
		        		continue; 
		        	default:
		        		throw new RuntimeException("Unhandled parse results!"); 
		        	} 
	        	} 
	        }  
	    } catch (Exception e) {
	        e.printStackTrace();
	    }
	}
	
	public static void main(String[] args) {
		
		TheGroupColumns.add(JoyGroupColumn);
		TheGroupColumns.add(GentleGroupColumn);
		TheGroupColumns.add(PeaceGroupColumn);
		TheGroupColumns.add(PeaceGroupColumn2);
		TheGroupColumns.add(PatienceGroupColumn);
		TheGroupColumns.add(AbundanceGroupColumn);
		TheGroupColumns.add(SelControlGroupColumn);
		TheGroupColumns.add(LoveGroupColumn);
		TheGroupColumns.add(KindnessGroupColumn);
		TheGroupColumns.add(FaithGroupColumn);
		TheGroupColumns.add(CantoneseGroupColumn);
		TheGroupColumns.add(GoodnessGroupColumn);  
		readFromExcelFile(args);
	}
	
	
	/*
	 * This method populates the cell groups in the Excel sheet.
	 */
	private static void populateCellGroups(Sheet aSheet, Hashtable<Integer, CellGroup> groupTable) {
		Row groupRow = aSheet.getRow(CellGroupRow);
		for (int i = 0 ; i < NumberOfCellGroups; i++) {
			Cell aCell = groupRow.getCell(i);
			String groupName = aCell.getStringCellValue();
			if (groupName != null && groupName.trim().length() > 0) {
				CellGroup aGroup = new CellGroup();
				aGroup.setGroupName(groupName);
				aGroup.setGroupNumber(TotalGroupCount);
				groupTable.put(TotalGroupCount,  aGroup);
				System.out.println("Group name = " + groupName); 
				TotalGroupCount++;
			}
		} 		
	}
	
	
	
	/**
	 * Parse the contents of a cell.
	 * When the cell has two names - put them into one family and continue to the next row for phone number.
	 * When the cell has one name - try to find out whether there is a phone number right next to it - search the digit.
	 * If the cell only has one name and no number - try to find out whether the phone number is in the next row or not.
	 * @param aCell
	 * @param aFamily
	 */
	private static int parseCell(Sheet aSheet, Cell aCell, Family aFamily) {
		int result = Undefined;
		String names = aCell.getStringCellValue();
		if (names != null && names.trim().length() == 0) {
			return result;
		}
		if (names.matches(".*\\d+.*")) {
			// this is the case of a single person
			// there is a digit in this string 
			String[] nameAndPhoneString = splitOnFirstDigit(names);
			Person aPerson = new Person();
			aPerson.setChineseName(nameAndPhoneString[0]);
			PhoneNumber aPhoneNumber = PhoneNumber.parsePhoneNumber(nameAndPhoneString[1]);
			aFamily.setPerson1(aPerson);
			aFamily.setPhoneNumber(aPhoneNumber); 
			result = SingleMemberWithPhone;
		} else {
			// the name parsing needs human eyes to figure out - i don't know whether the
			// first person's name has two characters or three characters
			
			String[] bothNames = names.split(" ");
			if (bothNames.length == 2) {
				if (bothNames[0].trim().length() > 1) {
					// first person has at least 2 characters
					Person person1 = new Person();
					person1.setChineseName(bothNames[0]);
					Person person2 = new Person();
					person2.setChineseName(bothNames[1]);
					Row phoneRow = aSheet.getRow(aCell.getRow().getRowNum() + 1);
					Cell phoneCell = phoneRow.getCell(aCell.getColumnIndex());
					PhoneNumber aPhoneNumber = PhoneNumber.parsePhoneNumber(phoneCell.getStringCellValue());
					aFamily.setPerson1(person1);
					aFamily.setPerson2(person2);
					aFamily.setPhoneNumber(aPhoneNumber);
					result = BothMemberWithPhone;
				}  
			}  
		}
		return result;		
	}
	
	/*
	 * Some cell contents have single name followed by a phone number. 
	 * This method finds out such and splits the contents into two elements.
	 */
	private static String[] splitOnFirstDigit(String inString) {
		String[] arr = inString.split("\\d+", 2);
	    String pt1 = arr[0].trim();
	    String pt2 = inString.substring(pt1.length() + 1).trim();
	    String [] output = new String[2];
	    output[0] = pt1;
	    output[1] = pt1;
	    return output;
	} 
	
	/*
	 * 
	 * 
	Old code for reference only.....
	
	//Iterate through each rows one by one
	        Iterator<Row> rowIterator = sheet.iterator();
	        while (rowIterator.hasNext())
	        {
	            Row row = rowIterator.next();
	            //For each row, iterate through all the columns
	            Iterator<Cell> cellIterator = row.cellIterator();

	            while (cellIterator.hasNext()) 
	            {
	                Cell cell = cellIterator.next();
	                System.out.print("Cell[" + cell.getRowIndex() + ", " + cell.getColumnIndex() + "] = ");
	                //Check the cell type and format accordingly
	                switch (cell.getCellType()) 
	                {
	                    case Cell.CELL_TYPE_NUMERIC:
	                        System.out.print(cell.getNumericCellValue() + "\t");
	                        break;
	                    case Cell.CELL_TYPE_STRING:
	                        System.out.print(cell.getStringCellValue() + "\t");
	                        break;
	                }
	            }
	            System.out.println("");
	        }
	
	 */
	 
}
