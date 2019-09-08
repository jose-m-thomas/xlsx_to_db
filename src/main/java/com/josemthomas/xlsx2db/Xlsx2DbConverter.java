package com.josemthomas.xlsx2db;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.sql.SQLType;
import java.sql.Types;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.log4j.BasicConfigurator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.fasterxml.jackson.databind.ObjectMapper;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class Xlsx2DbConverter {

	private File excelFile;
	private Connection connection;
	Map<String, Schema> schemas;

	static {
		BasicConfigurator.configure();
	}

	Xlsx2DbConverter(File excelFile, Connection connection) throws SQLException {
		this.connection = connection;
		connection.setAutoCommit(true);
		this.excelFile = excelFile;
	}

	public void convert(boolean replaceData, List<String> exclusionSheets) throws IOException, SQLException {
		long start=Calendar.getInstance().getTimeInMillis();

		FileInputStream file = new FileInputStream(excelFile);

		// Create Workbook instance holding reference to .xlsx file

		buildDbSchema(replaceData, exclusionSheets);
		populatedData();

		file.close();
		long end=Calendar.getInstance().getTimeInMillis();
		log.info("Overall took "+(end-start)+" milli Seconds");

	}

	private void populatedData() throws SQLException, IOException {

		FileInputStream file = new FileInputStream(excelFile);
		XSSFWorkbook workBook = new XSSFWorkbook(file);
		ObjectMapper mapper = new ObjectMapper();

		for (int i = 0; i < workBook.getNumberOfSheets(); i++) {
			long start=Calendar.getInstance().getTimeInMillis();
			XSSFSheet sheet = workBook.getSheetAt(i);
			Schema schema = schemas.get(sheet.getSheetName());
			String insertSql = getInsertQuery(schema);
			int count = 0;
			for (Row row : sheet) {

				if (count < 1) {
					count++;
					continue;
				}
				int ix = 0;
				List<String> paramList = new ArrayList<String>();
				for (Cell cell : row) {
					if (cell.getColumnIndex() > ix) {
						for (; ix < cell.getColumnIndex();) {
							paramList.add(null);
							ix++;
						}
					}
					String val = getCellValueAsString(cell);
					paramList.add(val);
					ix++;
					// schema.add(col);
				}
				for (; ix < schema.getColList().size();) {
					paramList.add(null);
					ix++;
				}

				executeSql(insertSql, paramList);

				count++;
			}
			// schemas.put(sheet.getSheetName(), schema);
			long end=Calendar.getInstance().getTimeInMillis();
			log.info(sheet.getSheetName()+"[ "+count+" Rows] got populated in "+(end-start)+" milli Seconds");

		}
	}

	private String getInsertQuery(Schema schema) {
		StringBuilder s = new StringBuilder();
		s.append("insert into " + schema.getDbName() + " (");
		String coma = "";
		int ix = 0;
		for (Col col : schema.getColList()) {
			s.append(coma + col.getDbName());
			coma = ",";
			ix++;
		}
		s.append(") values (");
		coma = "";
		for (int i = 0; i < ix; i++) {
			s.append(coma + "?");
			coma = ",";
		}
		s.append(")");
		// TODO Auto-generated method stub
		return s.toString();
	}

	private void buildDbSchema(boolean replaceData, List<String> exclusionSheets) throws IOException, SQLException {
		long start = Calendar.getInstance().getTimeInMillis();
		FileInputStream file = new FileInputStream(excelFile);
		XSSFWorkbook workBook = new XSSFWorkbook(file);
		log.info("Number of Sheets" + workBook.getNumberOfSheets() + "");
		ObjectMapper mapper = new ObjectMapper();
		schemas = new HashMap<String, Schema>();
		for (int i = 0; i < workBook.getNumberOfSheets(); i++) {
			XSSFSheet sheet = workBook.getSheetAt(i);

			Schema schema = new Schema();
			schema.setSheetName(sheet.getSheetName());
			schema.setDbName(sheet.getSheetName());
			int count = 0;
			for (Row row : sheet) {
				if (count > 0)
					break;
				int ix = 0;
				for (Cell cell : row) {
					Col col = new Col();
					String val = getCellValueAsString(cell);
					col.setColIndex(ix);
					col.setDbName(getDbColName(schema, val));
					col.setLength(getLength(sheet, ix));
					col.setType(cell.getCellType() + "");
					col.setDbType(getDBType(cell.getCellType() + ""));
					ix++;
					schema.add(col);
				}
				count++;
			}
			schemas.put(sheet.getSheetName(), schema);

			// for(sheet.get)
		}
		// log.info(mapper.writerWithDefaultPrettyPrinter().writeValueAsString(schemas));
		file.close();
		long end = Calendar.getInstance().getTimeInMillis();
		log.info("Schema Building took:" + (end - start) + " Milli Seconds");

		for (Schema schema : schemas.values()) {
			StringBuilder s = new StringBuilder();
			s.append("CREATE table " + schema.getDbName() + "(");
			String coma = "";
			for (Col col : schema.getColList()) {
				s.append(coma + "\n    " + col.getDbName() + " " + col.getDbType() + "(" + col.getLength() + ")");
				;
				coma = ",";
			}
			s.append("\n)");
			// log.info(s.toString());

			executeSql(s.toString(), null);
		}

	}

	private void executeSql(String sql, List<String> paramList) throws SQLException {
		PreparedStatement stmt = null;
		try {
			stmt = connection.prepareStatement(sql);
			if (paramList != null) {
				for (int i = 1; i <= paramList.size(); i++) {
					if (paramList.get(i - 1) != null) {
						stmt.setString(i, paramList.get(i - 1));
					} else {
						stmt.setNull(i, Types.VARCHAR);
					}
				}
			}
			int x = stmt.executeUpdate();

		} catch (SQLException e) {
			System.out.println(sql + paramList);
			e.printStackTrace();

			throw e;
		} finally {
			if (stmt != null) {
				stmt.close();
			}
		}

	}

	private String getDBType(String type) {
		// if (type.equalsIgnoreCase("STRING")) {
		return "VARCHAR";
		// }
		// TO BE ENHANCED LATER for other type , now Everything String
		// return null;
	}

	private String getDbColName(Schema schema, String val) {
		// TODO Auto-generated method stub
		val = val.replaceAll(" ", "_");
		return val;
	}

	private int getLength(XSSFSheet sheet, int ix) {
		int length = 0;
		for (Row row : sheet) {
			Cell cell = row.getCell(ix);
			String str = getCellValueAsString(cell);
			if (str.length() > length) {
				length = str.length();
			}
		}
		return length;
	}

	private static String getCellValueAsString(Cell cell) {
		if (cell != null) {
			switch (cell.getCellType()) {
			case BOOLEAN:
				return cell.getBooleanCellValue() + "";
			case STRING:
				return cell.getRichStringCellValue() + "";
			case NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					return cell.getDateCellValue() + "";
				} else {
					return cell.getNumericCellValue() + "";
				}
			case FORMULA:
				return cell.getCellFormula() + "";
			case BLANK:
				return ("");
			default:
				return ("");
			}
		}
		return "";

	}

}
