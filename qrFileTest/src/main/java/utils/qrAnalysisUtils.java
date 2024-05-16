package utils;

import com.google.zxing.BinaryBitmap;
import com.google.zxing.MultiFormatReader;
import com.google.zxing.NotFoundException;
import com.google.zxing.Result;
import com.google.zxing.client.j2se.BufferedImageLuminanceSource;
import com.google.zxing.common.HybridBinarizer;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;

public class qrAnalysisUtils {

    public static String qrToContent(File file){

        BufferedImage image = null;
        try {
            image = ImageIO.read(file);
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException("File read failure");
        }

        BufferedImageLuminanceSource source = new BufferedImageLuminanceSource(image);
        HybridBinarizer hybridBinarizer = new HybridBinarizer(source);
        BinaryBitmap bitmap = new BinaryBitmap(hybridBinarizer);

        MultiFormatReader multiFormatReader = new MultiFormatReader();
        Result result;
        try {
            result = multiFormatReader.decode(bitmap);
            return result.getText();
        } catch (NotFoundException e) {
            throw new RuntimeException("Qr code cannot be decoded");
        }
    }

}
