package com.doc2html.bean;

import java.util.ArrayList;
import java.util.List;

public class HtmlTableTr {

	private List<HtmlTableTd> tdList = new ArrayList<HtmlTableTd>();

	public HtmlTableTd getTd(int index) {
		return tdList.get(index);
	}

	public void addTd(HtmlTableTd td) {
		tdList.add(td);
	}

	public List<HtmlTableTd> getTdList() {
		return tdList;
	}

}
