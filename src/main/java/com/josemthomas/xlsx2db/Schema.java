package com.josemthomas.xlsx2db;

import java.util.ArrayList;
import java.util.List;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class Schema {
	String sheetName;
	String dbName;
	List<Col> colList;

	public void add(Col col) {
		if (colList == null) {
			colList = new ArrayList<Col>();
		}
		colList.add(col);
	}
}
