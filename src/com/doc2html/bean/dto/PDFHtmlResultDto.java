package com.doc2html.bean.dto;

import java.util.List;

public class PDFHtmlResultDto extends DocHtmlDto {

	@Override
	public String getHtml() {
		List<String> imageList = getImageList();
		if (imageList != null && !imageList.isEmpty()) {
			StringBuilder result = new StringBuilder();
			for (String imagePath : imageList) {
				result.append("<img src='" + imagePath + "'/>");
			}
			return result.toString();
		}
		return null;
	}
}
