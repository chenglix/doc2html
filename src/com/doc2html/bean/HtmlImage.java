package com.doc2html.bean;

import lombok.Getter;
import lombok.Setter;

@Setter
@Getter
public class HtmlImage {

	private String path;
	private int top;
	private int left;
	private int width;
	private int height;

	@Override
	public String toString() {
		StringBuilder result = new StringBuilder();
		result.append("<div src='").append(path).append("' ");
		result.append("style='background:#000000;position:absolute;");
		result.append("width:").append(width).append("px;");
		result.append("height:").append(height).append("px;");
		result.append("top:").append(top).append("px;");
		result.append("left:").append(left).append("px;");
		result.append("'></div>");
		return result.toString();
	}
}
