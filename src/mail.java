import operation.CreateExcel;

/**
 * 
 * @author n_sekiya
 */
public class mail  {

	
	public static void main(String[] args) {
		final CreateExcel ex1 = new CreateExcel("path");
		final CreateExcel ex2 = new CreateExcel("path");
		final String val1 = null;
		final String val2 = "hogehoge";
		ex1.addValue(val1);
		ex1.addValue(val2);
	}
}
