package com.josemthomas.xlsx2db;

import java.util.List;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class Col {
	int colIndex;
	String dbName;
	String dbType;
	String type;
	int length;
}
