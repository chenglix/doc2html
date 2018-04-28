package com.doc2html.bean.dto;

import java.util.List;

import lombok.Getter;
import lombok.Setter;

@Setter
@Getter
public class ExcelHtmlResultDto extends DocHtmlDto {

	private List<ExcelSheetDto> sheetList;
	private Integer activeSheetIndex;
}
