import java.io.*;
import java.util.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

abstract class  Table {
	public abstract void setValues(int column, Object value);
	//public boolean isHeadValid(Table head);
	public abstract boolean isDataValid();
}
/*
 * indicate one entry of a worksheet
 */
class Table1 extends Table {
	public String serNum;//serial number--0
	public String contNum;//contract number--1
	public String borrNum;//borrower number--2
	public String borrName;//borrower name--3
	public int time;//-2^31~2^31-1--4
	public int transAmount;//transaction amount--5
	public String curType;//currency type--6
	public String cautionMoney;//保证金代码--7
	public String cautionMoneyAccount;//保证金主账户--8
	public String paymentAcount;//支款账户--9
	public String shouxiAcount;//收息账户--10
	public String agencyNum;//经销商编码--11
	public String agencyName;//经销商名称--12
	public String processMark;//处理标识--13
	public String summary;//概要--14
	
	public Table1() {
		
	}
	
	//setters and getters
	@Override
	public void setValues(int columnIndex, Object value) {
		switch(columnIndex) {
		case 0:
			serNum = (String) value;
			break;
		case 1:
			contNum = (String) value;
			break;
		case 2:
			borrNum = (String) value;
			break;
		case 3:
			borrName = (String) value;
			break;
		case 4:
			time = (Integer) value;
			break;
		case 5:
			transAmount = (Integer) value;
			break;
		case 6:
			curType = (String) value;
			break;
		case 7:
			cautionMoney = (String) value;
			break;
		case 8:
			cautionMoneyAccount = (String) value;
			break;
		case 9:
			paymentAcount = (String) value;
			break;
		case 10:
			shouxiAcount = (String) value;
			break;
		case 11:
			agencyNum = (String) value;
			break;
		case 12:
			agencyName = (String) value;
			break;
		case 13:
			processMark = (String) value;
			break;	
		case 14:
			summary = (String) value;
			break;
		default://可能会产生越界错误
				
		}
	}
	
	@Override
	public boolean isDataValid() {
		if(null == serNum || null == borrNum  || null == (Integer)transAmount  
				|| null == cautionMoneyAccount  || null == summary)
			return false;
		return true;
	}
	
}


public class ExcelManager {
	
	private static Object getCellValue(Cell cell) {
	    switch (cell.getCellType()) {
    	case Cell.CELL_TYPE_STRING:
    		return cell.getStringCellValue();
 
    	case Cell.CELL_TYPE_BOOLEAN:
    		return cell.getBooleanCellValue();
 
    	case Cell.CELL_TYPE_NUMERIC:
    		return cell.getNumericCellValue();    		
	    }
	    return null;
	}
	/**
	 * 
	 * @param excelFilePath excel文件路径
	 * @param operation 操作类型
	 * @return
	 */
	public List<Table> excelRead(String excelFilePath, String operation) {
		List<Table> lt = new ArrayList<Table>();
		FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
		
		Workbook wb = new XSSFWorkbook(inputStream);
	    Sheet sh = wb.getSheetAt(0);
	    Iterator<Row> iterator = sh.iterator();
	    boolean isHead = true;
	   	    	    
	    while(iterator.hasNext()) {
	    	Row nextRow = iterator.next();
	    	Iterator<Cell> cellIterator = nextRow.cellIterator();
	    	//if(isHead == true) Object[] head = new Object[15];
	    	Table tb = TableFactory.getTable(operation); //used to store one data row of Table sheet
	    	
	    	while(cellIterator.hasNext()) {
	    		Cell nextCell = cellIterator.next();
	    		
	    		int columnIndex = nextCell.getColumnIndex();
	    		
	    		tb.setValues(columnIndex, ExcelManager.getCellValue(nextCell));
	    	}
	    	if(false == isHead && true == tb.isDataValid()) {
	    		lt.add(tb);
	    		isHead = false;
	    	}	    	
	    }
	    return lt;	    
	}
}
