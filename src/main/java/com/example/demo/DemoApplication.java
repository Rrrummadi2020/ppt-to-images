package com.example.demo;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xslf.usermodel.*;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.imageio.ImageIO;

import org.apache.poi.xslf.usermodel.*;

@SpringBootApplication
public class DemoApplication {

	public static void main(String[] args) {
		System.out.println("Hello World start");
		SpringApplication.run(DemoApplication.class, args);
		System.out.println("Hello World end");
		System.out.println("C:\\Users\\rrr23\\Downloads\\node_js_crash_course.pptx");
		// hello();
		kello();
	}

	public static void kello() {
		String pptxFilePath = "C:\\Users\\rrr23\\Downloads\\node_js_crash_course.pptx";
		String outputFolder = "D:\\hhuu\\";

		try {
			FileInputStream fis = new FileInputStream(pptxFilePath);
			XMLSlideShow ppt = new XMLSlideShow(fis);
			fis.close();

			// Create output directory if it doesn't exist
			File directory = new File(outputFolder);
			if (!directory.exists()) {
				directory.mkdir();
			}

			// Extract each slide as an image
			for (int i = 0; i < ppt.getSlides().size(); i++) {
				BufferedImage image = renderSlideAsImage(ppt, i);
				File outputFile = new File(outputFolder + File.separator + "slide" + (i + 1) + ".png");
				ImageIO.write(image, "png", outputFile);
			}

			System.out.println("Slides converted to images successfully.");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static BufferedImage renderSlideAsImage(XMLSlideShow ppt, int slideIndex) throws IOException {
		Dimension slideSize = ppt.getPageSize();
		XSLFSlide slide = ppt.getSlides().get(slideIndex);
		BufferedImage image = new BufferedImage(slideSize.width, slideSize.height, BufferedImage.TYPE_INT_RGB);
		Graphics2D graphics = image.createGraphics();
		graphics.setPaint(Color.white);
		graphics.fill(new Rectangle2D.Float(0, 0, slideSize.width, slideSize.height));
		slide.draw(graphics);
		graphics.dispose();
		return image;
	}

	public static void hello() {

		System.out.println("Hello World");
		try {
			// Provide the path to your PPTX file
			String pptxFile = "C:\\Users\\rrr23\\Downloads\\node_js_crash_course.pptx";

			// Open the PowerPoint file
			System.out.println("Opening PowerPoint file...");
			;
			FileInputStream fis = new FileInputStream(pptxFile);
			XMLSlideShow ppt = new XMLSlideShow(fis);
			fis.close();

			// Get slides from the PowerPoint presentation
			List<XSLFSlide> slides = ppt.getSlides();

			// Print content of each slide
			for (int i = 0; i < slides.size(); i++) {
				System.out.println("Slide " + (i + 1) + ":");
				System.out.println(slides.get(i).getTitle());
				System.out.println(slides.get(i).getPlaceholder(0)); // Get the text content, assuming it's in the first
																		// placeholder
				System.out.println();
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
