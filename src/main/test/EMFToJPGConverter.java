import org.freehep.graphicsio.emf.EMFInputStream;
import org.freehep.graphicsio.emf.EMFRenderer;

import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import javax.imageio.ImageIO;

public class EMFToJPGConverter {

    public static void main(String[] args) {

        String inputFilePath = "C:\\Users\\17600118628_98407\\IdeaProjects\\GoTest\\images\\image-1a7f4.emf";
        String outputFilePath = "C:\\Users\\17600118628_98407\\IdeaProjects\\GoTest\\images\\image-1a7f4.jpg";

        EMFInputStream inputStream = null;
        try {
            inputStream = new EMFInputStream(new FileInputStream(inputFilePath), EMFInputStream.DEFAULT_VERSION);
            System.out.println("height = " + inputStream.readHeader().getBounds().getHeight());
            System.out.println("widht = " + inputStream.readHeader().getBounds().getWidth());

            EMFRenderer emfRenderer = new EMFRenderer(inputStream);

            final int width = (int)emfRenderer.getSize().getWidth();
            final int height = (int)emfRenderer.getSize().getHeight();
            System.out.println("widht = " + width + " and height = " + height);
            final BufferedImage result = new BufferedImage(width*2, height*2, BufferedImage.TYPE_INT_RGB);

            Graphics2D g2 = (Graphics2D)result.createGraphics();
            // 绘制 EMF 内容
//            g2.setBackground() // 设置背景为白色
            g2.setBackground(Color.WHITE);
            g2.clearRect(0, 0, width*3, height*3); // 清空画布，防止内容超出

            // 缩放绘图
            g2.scale(2, 2);

            emfRenderer.paint(g2);

            // 释放资源
            g2.dispose();
            inputStream.close();


            File outputfile = new File(outputFilePath);
            ImageIO.write(result, "jpg", outputfile);

        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }
}