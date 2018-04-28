package com.doc2html;

import java.io.IOException;
import java.util.List;

import com.doc2html.bean.dto.DocHtmlDto;
import com.doc2html.bean.dto.ExcelHtmlResultDto;
import com.doc2html.bean.dto.ExcelSheetDto;

public class Test {

	public static void main(String[] args) throws IOException {
		testExcel();
	}

	private static void testWord() throws IOException {
		String filePath = "xxxxxxxx";
		Doc2Html doc2Html = new Word2Html();
		DocHtmlDto resultDto = doc2Html.doc2Html(filePath);
		if (resultDto != null) {
			System.out.println(resultDto.getHtml());
		} else {
			System.out.println("×ª»»Ê§°Ü");
		}
	}

	private static void testExcel() throws IOException {
		String filePath = "xxxxxx";
		Doc2Html doc2Html = new Excel2Html();
		ExcelHtmlResultDto resultDto = (ExcelHtmlResultDto) doc2Html.doc2Html(filePath);
		if (resultDto != null) {
			List<ExcelSheetDto> sheetList = resultDto.getSheetList();
			for (ExcelSheetDto sheetDto : sheetList) {
				System.out.println(sheetDto.getHtml());
			}
		} else {
			System.out.println("×ª»»Ê§°Ü");
		}
	}

	private static void testPDF() throws IOException {
		String filePath = "xxxxxxxxxxxxxx";
		Doc2Html doc2Html = new PDF2Html();
		DocHtmlDto resultDto = doc2Html.doc2Html(filePath);
		System.out.println(resultDto.getHtml());
	}

	private static void testPPT() throws IOException {
		String filePath = "xxxxxxxxxxxxxxxxxx";
		Doc2Html doc2Html = new PPT2Html();
		DocHtmlDto resultDto = doc2Html.doc2Html(filePath);
		System.out.println(resultDto.getHtml());
	}
}
