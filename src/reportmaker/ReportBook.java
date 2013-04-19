/**
 * 
 */
package reportmaker;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.ResultSet;
import java.sql.Statement;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;


/**
 * @author Maxim Surtaev
 *
 */
public class ReportBook {
	private SXSSFWorkbook reportBook;
	private Connection dbConnection;
	
	
	public ReportBook (int ParamRowAccessWindowSize) {
		reportBook=new SXSSFWorkbook(ParamRowAccessWindowSize);
		
	}
	public void connectToDB (String jdbcString, String dbUser, String dbUserPassword) throws SQLException {
		dbConnection=DriverManager.getConnection(jdbcString, dbUser, dbUserPassword);
			
	}
	public void addReportSheet (String ParamSheetName, String dbQuery) throws SQLException {
		Sheet sheet=reportBook.createSheet(ParamSheetName);
		Statement reportStatement=dbConnection.createStatement();
		ResultSet reportResultSet=reportStatement.executeQuery(dbQuery);
		ReportSheet reportSheet=new ReportSheet(sheet,reportResultSet);
		reportSheet.fillSheet();
	}
	

}
