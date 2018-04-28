package com.doc2html;

import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.imageio.ImageIO;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;

import com.doc2html.bean.dto.DocHtmlDto;
import com.doc2html.bean.dto.PDFHtmlResultDto;
import com.doc2html.config.Config;

public class PDF2Html implements Doc2Html {

	@Override
	public DocHtmlDto doc2Html(String filePath) throws IOException {
		PDDocument doc = PDDocument.load(new File(filePath));
		PDFRenderer renderer = new PDFRenderer(doc);
		int pageCount = doc.getNumberOfPages();
		List<String> imageList = new ArrayList<>(pageCount);
		String dir = Config.getDir(System.currentTimeMillis() + "");
		File folderFile = new File(dir);
		if (!folderFile.exists()) {
			folderFile.mkdirs();
		}
		for (int i = 0; i < pageCount; i++) {
			BufferedImage image = renderer.renderImageWithDPI(i, 200, ImageType.RGB);
			if (image.getWidth() > 1000) {
				image = contractImage(image, 1000);
			}
			String fileName = dir + "/" + i + ".jpg";
			ImageIO.write(image, "JPG", new File(fileName));
			imageList.add(fileName);
		}
		System.out.println("全部处理完毕！");

		PDFHtmlResultDto resultDto = new PDFHtmlResultDto();
		resultDto.setImageList(imageList);
		return resultDto;
	}

	private static BufferedImage contractImage(Image image, int width) {

		if (image == null) {
			return null;
		}
		int imageWidth = image.getWidth(null);
		int imageHeight = image.getHeight(null);

		if (imageWidth <= width) {
			width = imageWidth;
		}
		int height = imageHeight * width / imageWidth;

		Image newImage = image.getScaledInstance(width, height, Image.SCALE_SMOOTH);

		BufferedImage bImage = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
		Graphics2D gf = bImage.createGraphics();
		gf.dispose();
		gf = bImage.createGraphics();
		// 处理锯齿
		gf.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);
		gf.drawImage(newImage, 0, 0, width, height, null);
		gf.dispose();
		return bImage;

	}

}
