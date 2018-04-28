package com.doc2html;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import com.doc2html.bean.dto.DocHtmlDto;

public class Doc2HtmlFactory {

	private static Doc2Html excel2Html = new Excel2Html();
	private static Doc2Html pdf2Html = new PDF2Html();
	private static Doc2Html ppt2Html = new PPT2Html();
	private static Doc2Html word2Html = new Word2Html();

	public static Map<String, Doc2Html> instanceMap = new HashMap<String, Doc2Html>();
	static {
		instanceMap.put("doc", word2Html);
		instanceMap.put("docx", word2Html);
		instanceMap.put("xls", excel2Html);
		instanceMap.put("xlsx", excel2Html);
		instanceMap.put("ppt", ppt2Html);
		instanceMap.put("pptx", ppt2Html);
		instanceMap.put("pdf", pdf2Html);
	}

	public static Doc2Html getInstance(String filePath) {
		if (filePath.indexOf(".") >= 0) {
			String fileType = filePath.substring(filePath.lastIndexOf(".") + 1, filePath.length());
			Doc2Html doc2Html = instanceMap.get(fileType.toLowerCase());
			return doc2Html;
		}
		return null;
	}

	public static DocHtmlDto coverToHtml(String filePath) {
		Doc2Html doc2Html = getInstance(filePath);
		if (doc2Html != null) {
			try {
				return doc2Html.doc2Html(filePath);
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return null;
	}
}
