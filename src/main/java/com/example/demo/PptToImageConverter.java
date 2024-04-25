package com.example.demo;

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


public class PptToImageConverter {

    public static void main(String[] args) {
        String pptxFilePath = "presentation.pptx";
        String outputFolder = "output";

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
}
