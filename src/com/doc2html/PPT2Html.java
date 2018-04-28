package com.doc2html;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.RenderingHints;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.imageio.ImageIO;

import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFSlideShowImpl;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import com.doc2html.bean.dto.DocHtmlDto;
import com.doc2html.bean.dto.PPTHtmlResultDto;
import com.doc2html.config.Config;

public class PPT2Html implements Doc2Html {

	public DocHtmlDto doc2Html(String filePath) throws IOException {
		DocHtmlDto result = null;
		if (filePath.contains(".")) {
			String fileType = filePath.substring(filePath.lastIndexOf(".") + 1, filePath.length());
			if ("ppt".equals(fileType)) {
				result = toImage2003(filePath);
			} else if ("pptx".equals(fileType)) {
				result = toImage2007(filePath);
			}
		}
		System.out.println("转换完毕");
		return result;
	}

	private static PPTHtmlResultDto toImage2003(String filePath) throws IOException {

		String dir = Config.getDir(System.currentTimeMillis() + "");
		File f = new File(dir.toString());
		if (!f.exists()) {
			f.mkdirs();
		}

		HSLFSlideShow ppt = new HSLFSlideShow(new HSLFSlideShowImpl(filePath));
		Dimension pgSize = ppt.getPageSize();

		int width = pgSize.width;
		int height = pgSize.height;

		List<HSLFSlide> sliders = ppt.getSlides();
		int sliderSize = sliders.size();

		List<String> imageList = new ArrayList<String>();
		for (int i = 0; i < sliderSize; i++) {
			BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
			Graphics2D graphics = image.createGraphics();

			// 设置对线段的锯齿状边缘处理
			graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);

			graphics.drawImage(image.getScaledInstance(width, height, Image.SCALE_SMOOTH), 0, 0, null);

			graphics.setPaint(Color.white);
			graphics.fill(new Rectangle2D.Float(0, 0, width, height));

			ppt.getSlides().get(i).draw(graphics);

			// save the output
			String filename = dir + "/" + (i + 1) + ".jpg";
			FileOutputStream out = new FileOutputStream(filename);
			ImageIO.write(image, "jpg", out);
			out.close();
			imageList.add(filename);
		}
		ppt.close();
		PPTHtmlResultDto resultDto = new PPTHtmlResultDto();
		resultDto.setImageList(imageList);
		return resultDto;
	}

	private static PPTHtmlResultDto toImage2007(String filePath) throws IOException {

		String dir = Config.getDir(System.currentTimeMillis() + "");
		File f = new File(dir);
		if (!f.exists()) {
			f.mkdirs();
		}
		FileInputStream inp = new FileInputStream(filePath);
		XMLSlideShow ppt = new XMLSlideShow(inp);
		inp.close();
		Dimension pgSize = ppt.getPageSize();

		int width = pgSize.width;
		int height = pgSize.height;

		List<XSLFSlide> sliders = ppt.getSlides();
		int sliderSize = sliders.size();
		List<String> imageList = new ArrayList<String>(sliderSize);
		for (int i = 0; i < sliderSize; i++) {
			BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
			Graphics2D graphics = image.createGraphics();
			// 设置对线段的锯齿状边缘处理
			graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);

			graphics.drawImage(image.getScaledInstance(width, height, Image.SCALE_SMOOTH), 0, 0, null);
			graphics.setPaint(Color.white);
			graphics.fill(new Rectangle2D.Float(0, 0, width * 3, height * 3));

			ppt.getSlides().get(i).draw(graphics);

			// save the output
			String filename = dir + "/" + (i + 1) + ".jpg";
			FileOutputStream out = new FileOutputStream(filename);
			javax.imageio.ImageIO.write(image, "jpg", out);
			out.close();
			imageList.add(filename);
		}
		ppt.close();

		PPTHtmlResultDto resultDto = new PPTHtmlResultDto();
		resultDto.setImageList(imageList);
		return resultDto;
	}

}
