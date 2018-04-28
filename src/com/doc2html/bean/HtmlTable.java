package com.doc2html.bean;

import java.util.ArrayList;
import java.util.List;

public class HtmlTable {

	private List<HtmlTableTr> trList = new ArrayList<HtmlTableTr>();
	private List<HtmlImage> imageList = null;

	public void addTr(HtmlTableTr tr) {
		trList.add(tr);
	}

	public HtmlTableTr getTr(int index) {
		return trList.get(index);
	}

	public void setImageList(List<HtmlImage> imageList) {
		this.imageList = imageList;
	}

	@Override
	public String toString() {
		StringBuilder html = new StringBuilder();
		html.append("<table cellspacing='0' cellpadding='0' style='border-collapse:collapse;'>");
		for (HtmlTableTr tr : trList) {
			html.append("<tr>");
			List<HtmlTableTd> tdList = tr.getTdList();
			for (HtmlTableTd td : tdList) {
				if (td != null) {
					html.append(td.toString());
				}
			}
			html.append("</tr>");
		}
		html.append("</table>");
		if (imageList != null && !imageList.isEmpty()) {
			for (HtmlImage image : imageList) {
				html.append(image.toString());
			}
		}
		return html.toString();
	}

	public List<HtmlTableTr> getTrList() {
		return trList;
	}

}
