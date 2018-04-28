package com.doc2html.bean;

import lombok.Getter;
import lombok.Setter;

@Setter
@Getter
public class ExcelPictureData {

	private int row1;
	private int row2;

	private int col1;
	private int col2;

	private int col1X;
	private int col2X;

	private int row1Y;
	private int row2Y;

	private String path;
}
