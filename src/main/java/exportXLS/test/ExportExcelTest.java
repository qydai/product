package exportXLS.test;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.SimpleTimeZone;

import exportXLS.entity.UserInfo;
import exportXLS.util.ExcelUtil;



public class ExportExcelTest {

    public static void main(String[] args) throws Exception{
        String thisPath = ExportExcelTest.class.getResource("").getPath();
        String basePath = thisPath.substring(1,thisPath.indexOf("/exportXLS"));
        String filePath = basePath + "/5603d787N741c873e.jpg";
        List<UserInfo> users = new ArrayList<UserInfo>();
        UserInfo user = new UserInfo();
        user.setAddress("江西省丰城市");
        user.setId(1);
        user.setBirthDate(new Date());
        user.setIdCard("362202********6354");
        user.setQq("101***603");
        user.setSex(0);
        user.setName("戴某某");
        user.setPhone("183****4800");
        user.setPhoto(file2Byte(filePath));
        users.add(user);
        user = new UserInfo();
        user.setAddress("江西省丰城市");
        user.setId(2);
        user.setPhoto(file2Byte(filePath));
        user.setBirthDate(new Date());
        user.setIdCard("362202********6354");
        user.setQq("101***603");
        user.setSex(1);
        user.setPhone("186****3897");
        user.setName("王某某");
        users.add(user);
        user = new UserInfo();
        user.setAddress("江西省丰城市");
        user.setId(3);
        user.setPhoto(file2Byte(filePath));
        user.setBirthDate(new Date());
        user.setIdCard("362202********6354");
        user.setQq("101***603");
        user.setSex(0);
        user.setName("戴某某");
        user.setPhone("186***5177");
        users.add(user);
        user = new UserInfo();
        user.setAddress("江西省丰城市");
        user.setId(4);
        user.setPhoto(file2Byte(filePath));
        user.setBirthDate(new Date());
        user.setIdCard("362202********6354");
        user.setQq("101***603");
        user.setSex(1);
        user.setPhone("186****3897");
        user.setName("王某某");
        users.add(user);
        user = new UserInfo();
        user.setAddress("江西省丰城市");
        user.setId(5);
        user.setPhoto(file2Byte(filePath));
        user.setBirthDate(new Date());
        user.setIdCard("362202********6354");
        user.setQq("101***603");
        user.setSex(1);
        user.setPhone("186****3897");
        user.setName("王某某");
        users.add(user);
        user = new UserInfo();
        user.setAddress("江西省丰城市");
        user.setId(6);
        user.setPhoto(file2Byte(filePath));
        user.setBirthDate(new Date());
        user.setIdCard("362202********6354");
        user.setQq("101***603");
        user.setSex(1);
        user.setPhone("186****3897");
        user.setName("王某某");
        users.add(user);
        user = new UserInfo();
        user.setAddress("江西省丰城市");
        user.setId(7);
        user.setPhoto(file2Byte(filePath));
        user.setBirthDate(new Date());
        user.setIdCard("362202********6354");
        user.setQq("101***603");
        user.setSex(1);
        user.setPhone("186****3897");
        user.setName("王某某");
        users.add(user);
        user = new UserInfo();
        user.setAddress("江西省丰城市");
        user.setId(8);
        user.setPhoto(file2Byte(filePath));
        user.setBirthDate(new Date());
        user.setIdCard("362202********6354");
        user.setQq("101***603");
        user.setSex(1);
        user.setPhone("186****3897");
        user.setName("王某某");
        users.add(user);
        user = new UserInfo();
        user.setAddress("江西省丰城市");
        user.setId(9);
        user.setPhoto(file2Byte(filePath));
        user.setBirthDate(new Date());
        user.setIdCard("362202********6354");
        user.setQq("101***603");
        user.setSex(1);
        user.setPhone("186****3897");
        user.setName("王某某");
        users.add(user);
        user = new UserInfo();
        user.setAddress("江西省丰城市");
        user.setId(10);
        user.setBirthDate(new Date());
        user.setIdCard("362202********6354");
        user.setQq("101***603");
        user.setSex(1);
        user.setPhone("186****3897");
        user.setName("王某某");
        users.add(user);
        user = new UserInfo();
        user.setAddress("江西省丰城市");
        user.setId(11);
        user.setPhoto(file2Byte(filePath));
        user.setBirthDate(new Date());
        user.setIdCard("362202********6354");
        user.setQq("101***603");
        user.setSex(1);
        user.setPhone("186****3897");
        user.setName("王某某");
        users.add(user);
        user = new UserInfo();
        user.setAddress("江西省丰城市");
        user.setId(12);
        user.setPhoto(file2Byte(filePath));
        user.setBirthDate(new Date());
        user.setIdCard("362202********6354");
        user.setQq("101***603");
        user.setSex(1);
        user.setPhone("186****3897");
        user.setName("王某某");
        users.add(user);
        user = new UserInfo();
        user.setAddress("江西省丰城市");
        user.setId(13);
        user.setPhoto(file2Byte(filePath));
        user.setBirthDate(new Date());
        user.setIdCard("362202********6354");
        user.setQq("101***603");
        user.setSex(1);
        user.setPhone("186****3897");
        user.setName("王某某");
        users.add(user);
        user = new UserInfo();
        user.setAddress("江西省丰城市");
        user.setId(14);
        user.setPhoto(file2Byte(filePath));
        user.setBirthDate(new Date());
        user.setIdCard("362202********6354");
        user.setQq("101***603");
        user.setSex(1);
        user.setPhone("186****3897");
        user.setName("王某某");
        users.add(user);
        user = new UserInfo();
        user.setAddress("江西省丰城市");
        user.setId(15);
        user.setBirthDate(new Date());
        user.setIdCard("362202********6354");
        user.setQq("101***603");
        user.setSex(1);
        user.setPhone("186****3897");
        user.setName("王某某");
        users.add(user);
        user = new UserInfo();
        user.setAddress("江西省丰城市");
        user.setId(16);
        user.setPhoto(file2Byte(filePath));
        user.setBirthDate(new Date());
        user.setIdCard("362202********6354");
        user.setQq("101***603");
        user.setSex(1);
        user.setPhone("186****3897");
        user.setName("王某某");
        users.add(user);
        user = new UserInfo();
        user.setAddress("江西省丰城市");
        user.setId(17);
        user.setPhoto(file2Byte(filePath));
        user.setBirthDate(new Date());
        user.setIdCard("362202********6354");
        user.setQq("101***603");
        user.setSex(1);
        user.setPhone("186****3897");
        user.setName("王某某");
        users.add(user);
        user = new UserInfo();
        user.setAddress("江西省丰城市");
        user.setId(18);
        user.setPhoto(file2Byte(filePath));
        user.setBirthDate(new Date());
        user.setIdCard("362202********6354");
        user.setQq("101***603");
        user.setSex(1);
        user.setPhone("186****3897");
        user.setName("王某某");
        users.add(user);
        ExcelUtil util = new ExcelUtil();
        try {
            util.exportExcel(users,UserInfo.class);
        } catch (Exception e) {
            e.printStackTrace();
        }

    }
    public static byte[] file2Byte(String filePath){
        byte[] buffer = null;
        FileInputStream fis = null;
        ByteArrayOutputStream out = null;
        try{
            File file = new File(filePath);
            fis = new FileInputStream(file);
            out = new ByteArrayOutputStream();
            byte[] b = new byte[1024];
            int n = 0;
            while ((n = fis.read(b))!=-1){
                out.write(b,0,n);
            }
            buffer = out.toByteArray();
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            try {
                if (fis != null) {
                    fis.close();
                }
                if (out != null) {
                    out.close();
                }
            }catch (IOException e){
                e.printStackTrace();
            }
        }
        return buffer;
    }
    public static File byte2file(byte[] buffer){
        File file = null;
        FileOutputStream out = null;
        try {
            file = new File("D:/temple/temp"+new SimpleDateFormat("yyyyMMddHHmmssS").format(new Date())+".jpg");
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
