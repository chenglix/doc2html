package com.doc2html.bean;

import lombok.Getter;
import lombok.Setter;

@Setter
@Getter
public class HtmlTableTd {

	private String text;
	private Integer width;
	private Integer height;
	private Integer colspan;
	private Integer rowspan;
	private String background;
	private String valign;
	private String align;

	private ExcelFontInfo font;

	@Override
	public String toString() {
		StringBuilder str = new StringBuilder();
		str.append("<td style='");
		str.append("border:1px solid #000000;");
		if (background != null) {
			str.append("background:" + background + ";");
		}
		str.append("' ");
		if (width != null) {
			str.append(" width='").append(width).append("px' ");
		}
		if (height != null) {
			str.append(" height='").append(height).append("px' ");
		}
		if (rowspan != null) {
			str.append(" rowspan='").append(rowspan).append("' ");
		}
		if (colspan != null) {
			str.append(" colspan='").append(colspan).append("' ");
		}
		if (align != null) {
			str.append(" align='").append(align).append("' ");
		}
		if (valign != null) {
			str.append(" valign='").append(valign).append("' ");
		}
		str.append(">");
		if (text != null) {
			text = text.replace("\n", "<br>").replace(" ", "&nbsp;");
			if (font != null) {
				str.append("<span style='display:inline-block;overflow:hidden;");
				if (width != null) {
					str.append("width:").append(width).append("px;");
				}
				if (font.getFontSize() != null) {
					str.append("font-size:").append(font.getFontSize()).append("px;");
				}
				if (font.getFontFamily() != null) {
					str.append("font-family:").append(font.getFontFamily()).append(";");
				}
				if (font.getColor() != null) {
					str.append("color:").append(font.getColor()).append(";");
				}
				str.append("'>");
				str.append(text);
				str.append("</span>");
			} else {
				str.append(text);
			}
		} else {
			str.append("&nbsp;");
		}
		str.append("</td>");
		return str.toString();
	}

}
