package com.marzaha.pptx;


import android.content.Context;
import android.graphics.Bitmap;
import android.graphics.Canvas;
import android.graphics.Color;
import android.graphics.Paint;
import android.os.Handler;

import net.pbdavey.awt.Graphics2D;

import org.apache.commons.io.FileUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import java.io.File;
import java.io.FileOutputStream;
import java.util.concurrent.atomic.AtomicBoolean;

import and.awt.Dimension;

public class PointParser {

    private Context context;
    private String filePath;

    private int slideCount = 0;
    private XSLFSlide[] slide;
    private Dimension pgsize;
    private XMLSlideShow pptx = null;

    public String htmlPath ;

    public PointParser(Context context, String path) {
        this.filePath = path;
        this.context = context;
    }

    public void pptxToHtml(Handler handler) {
        try {
            // 初始化
            pptx = new XMLSlideShow(OPCPackage.open(filePath, PackageAccess.READ));
            pgsize = pptx.getPageSize();
            slide = pptx.getSlides();
            slideCount = slide.length;

            String imageHtml = "";

            for (int position = 0; position < slideCount; position++) {
                Bitmap bmp = Bitmap.createBitmap((int) pgsize.getWidth(), (int) pgsize.getHeight(), Bitmap.Config.RGB_565);
                Canvas canvas = new Canvas(bmp);
                Paint paint = new Paint();
                paint.setColor(Color.WHITE);
                paint.setFlags(Paint.ANTI_ALIAS_FLAG);
                canvas.drawPaint(paint);
                Graphics2D graphics2d = new Graphics2D(canvas);
                AtomicBoolean isCanceled = new AtomicBoolean(false);

                slide[position].draw(graphics2d, isCanceled, handler, position);

                String imageName = System.currentTimeMillis() + ".png";
                String imagePath = context.getCacheDir() + File.separator + imageName;
                File imageFile = new File(imagePath);
                if (!imageFile.exists()) {
                    imageFile.createNewFile();
                }
                FileOutputStream outputStream = new FileOutputStream(imageFile);
                bmp.compress(Bitmap.CompressFormat.PNG, 80, outputStream);
                outputStream.flush();

                imageHtml += "<img src=\'"
                        + "file:///" + imagePath
                        + "\' style=\'vertical-align:text-bottom; \' border='1'><br><br>";
            }
            String pptxHtml = "<html><head><META http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"></head><body>"
                    + imageHtml + "</body></html>";
            String pptxName = filePath.substring(filePath.lastIndexOf("/") + 1);
            String htmlName = pptxName.replace("pptx", "html");
            htmlPath = context.getCacheDir() + File.separator + htmlName;
            FileUtils.writeStringToFile(new File(htmlPath), pptxHtml, "utf-8");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


}
