package operation;

import java.sql.Date;
import java.util.HashMap;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateExcel {

	
	final Workbook wb;
	final Sheet sheet1;
	
	private Integer _colNumber;
	
//	public CreateExcel(final String path) {
//		// xlsの場合はこちらを有効化
//		wb = new HSSFWorkbook();
//	}

	public CreateExcel(final String path) {
		//xlsxの場合はこちらを有効化
		wb = new XSSFWorkbook();
		sheet1 = wb.createSheet("シート名");
	}

	private void test(){
		final HashMap map = new HashMap(); 
		for (Sheet sheet : wb ) {
			 map.put(1, sheet);
		 }
	}
	
	
	private void createRow() {
		
	}
	
	private void addValeu() {
		if (_colNumber == null) {
			sheet1.createRow(0);
			_colNumber = 1;
		} else {
			sheet1.createRow(_colNumber);
			_colNumber += 1;
		}
	}
	
	/**
	 * String型の出力処理
	 * @param value
	 */
	public void addValue(final String value) {
		// TODO
	}
	
	/**
	 * Integer型の出力処理
	 * @param value
	 */
	public void addValue(final Integer value) {
		// TODO
	}
	
	/**
	 * Boolean型の出力処理
	 * @param value
	 */
	public void addValue(final Boolean value) {
		// TODO
	}
	
	/**
	 * Date型の出力処理
	 * @param value
	 */
	public void addValue(final Date value) {
		// TODO
	}
	
}
