/**
 * 
 */
package reportmaker;

/**
 * @author Maxim Surtaev
 *
 */

import java.sql.ResultSet;
import java.sql.SQLException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ReportSheet {

	/**
	 * Excel can not work with more then 1000000 rows
	 */
	final int DEFAULT_MAX_ROW_COUNT = 1000000;
	private ResultSet reportResultSet;
	private Sheet reportSheet;

	/**
	 * @param ParamSheet
	 *            Sheet for inserting data, must be created
	 * @param ParamResultSet
	 *            ResultSet the Set must be opened
	 */
	public ReportSheet(Sheet ParamSheet, ResultSet ParamResultSet) {
		reportResultSet = ParamResultSet;
		reportSheet = ParamSheet;
	}

	/**
	 * @throws SQLException
	 */
	public void fillSheet() throws SQLException {
		Row row = reportSheet.createRow(0);
		Cell cell = null;
		/**
		 * Setting head of report
		 */
		for (int i = 0; i < reportResultSet.getMetaData().getColumnCount(); i++) {
			cell = row.createCell(i);
			cell.setCellValue(reportResultSet.getMetaData().getColumnName(i + 1));
		}
		reportResultSet.first();
		/**
		 * Filling data
		 */
		for (int i = 1; reportResultSet.next() && i <= DEFAULT_MAX_ROW_COUNT; i++) {
			row = reportSheet.createRow(i);
			for (int j = 0; j < reportResultSet.getMetaData().getColumnCount(); j++) {
				cell = row.createCell(j);
				cell.setCellValue(reportResultSet.getString(j + 1));
			}
		}
	}

	public Sheet getSheet() {
		return reportSheet;
	}

	public void clearSheet() {
		reportSheet = null;
	}
}
