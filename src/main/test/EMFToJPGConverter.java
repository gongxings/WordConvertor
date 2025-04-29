import org.freehep.graphicsio.emf.EMFInputStream;
import org.freehep.graphicsio.emf.EMFRenderer;

import java.awt.*;
import java.awt.geom.AffineTransform;
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

//            final int width = (int)emfRenderer.getSize().getWidth();
//            final int height = (int)emfRenderer.getSize().getHeight();
//            System.out.println("widht = " + width + " and height = " + height);
            Rectangle paddedBounds = new Rectangle(
                    inputStream.readHeader().getBounds().x-50,
                    inputStream.readHeader().getBounds().y-50,
                    inputStream.readHeader().getBounds().width + 50,
                    inputStream.readHeader().getBounds().height + 50
            );

            final BufferedImage result = new BufferedImage(paddedBounds.width, paddedBounds.height, BufferedImage.TYPE_INT_RGB);

            Graphics2D g2 = result.createGraphics();
            // 可选：设置抗锯齿等参数
            g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            g2.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);


            // 绘制 EMF 内容
//            g2.setBackground() // 设置背景为白色
            g2.setTransform(new AffineTransform());
            g2.setBackground(Color.WHITE);
            g2.fillRect(0, 0, result.getWidth(), result.getHeight());
//            g2.clearRect(paddedBounds.x,paddedBounds.y,paddedBounds.width,paddedBounds.height); // 清空画布，防止内容超出
            // 缩放绘图
            g2.scale(1, 1);

            emfRenderer.paint(g2);

            // 释放资源
            g2.dispose();
            inputStream.close();


            File outputfile = new File(outputFilePath);
            ImageIO.write(result, "jpg", outputfile);

        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }
}
