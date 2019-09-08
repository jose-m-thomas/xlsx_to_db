package com.josemthomas.xlsx2db;

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

import org.junit.Before;
import org.junit.Test;

/**
 * Unit test for simple App.
 */
public class AppTest {

	Connection conn;

	@Before
	public void setup() throws SQLException, ClassNotFoundException {
		// Class.forName("org.hsqldb.jdbcDriver");

		this.conn = DriverManager.getConnection("jdbc:h2:mem:testdb", "sa", null);
	}

	@Test
	public void testFrameWork() throws IOException, SQLException {
		ClassLoader classLoader = getClass().getClassLoader();
		File file = new File(classLoader.getResource("test.xlsx").getFile());
		Xlsx2DbConverter converter = new Xlsx2DbConverter(file, conn);
		converter.convert(true, null);
	}
}
