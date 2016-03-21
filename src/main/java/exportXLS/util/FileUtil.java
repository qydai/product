package exportXLS.util;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Created by qydai on 2016/3/11.
 */
public class FileUtil {
    /**
     * 将一byte数组转换成图片
     * @param buffer byte数组
     * @return 转换完成的图片
     */
    public static File byte2file(byte[] buffer,String fileType){
        File file = null;
        FileOutputStream out = null;
        try {
            file = new File("D:/temple/temp"+new SimpleDateFormat("yyyyMMddHHmmssS").format(new Date())+"."+fileType);
            File parentFile = file.getParentFile();
            if(!parentFile.exists()){
                parentFile.mkdirs();
            }
            if(!file.exists()){
                file.createNewFile();
            }
            out = new FileOutputStream(file);
            out.write(buffer);
            out.close();
        }catch (IOException e){
            e.printStackTrace();
        }
        return file;
    }
}
