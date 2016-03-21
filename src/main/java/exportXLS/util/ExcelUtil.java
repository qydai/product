package exportXLS.util;

import java.awt.image.BufferedImage;
import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Date;
import java.util.List;

import exportXLS.exception.ExportExcelException;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Font;

import exportXLS.annotation.Column;
import exportXLS.annotation.Xls;

import javax.imageio.ImageIO;


public class ExcelUtil {
    private String sheetName = "";
    private final static String PART_HEAD = "head";
    private final static String PART_BODY = "body";
    private HSSFWorkbook workbook;
    private final static Integer MAX_PAGE_NUM = 43000;
    private String fileName;
    private String imgType;

    /**
     * 导出并且生成 excel 文件
     *
     * @param entityList  数据对象的集合
     * @param entityClass 反射的类对象
     */
    public <T> void exportExcel(List<T> entityList, Class<T> entityClass) throws ExportExcelException{
        //获取反射类对象中的字段信息
        Field[] fields = entityClass.getDeclaredFields();
        //获取有注解Column的字段并根据Column.sortNo排序
        List<Field> fieldList = getExportFiledSort(fields);
        this.getWorkbook();
        int pageCount = 0;
        Xls excelClass = entityClass.getAnnotation(Xls.class);
        String sheetName = fileName =excelClass.name();
        int size = entityList.size();
        try {
            //创建工作页的行，根据list中的对象创建数据行，从滴1行开始
            if (size > 0) {
                pageCount = (size / MAX_PAGE_NUM) + 1;
                HSSFPatriarch patriarch = null;
                for (int i = 0; i < pageCount; i++) {
                    this.sheetName = sheetName + (i + 1);
                    patriarch = this.workbook.createSheet(this.sheetName).createDrawingPatriarch();
                    this.createSheetHeaderRow(fieldList);
                    for (int rowNo = 1, listNo = i * MAX_PAGE_NUM; listNo < (((i + 1) == pageCount) ? (i * MAX_PAGE_NUM) + (size % MAX_PAGE_NUM) : (i + 1) * MAX_PAGE_NUM); rowNo++, listNo++) {
                        this.createSheetBodyRow(rowNo, entityList.get(listNo), fieldList,patriarch);
                    }
                }
            }
            createExcelFile("D:/");
        }catch (Exception e){
            throw new ExportExcelException(e);
        }

    }

    /**
     * 获取有注解Column的字段并根据Column.sortNo排序
     * @param fields 对象的所有字段
     * @return 对象中所有注解为Column的字段信息
     */
    private List<Field> getExportFiledSort(Field[] fields) {
        List<Field> fieldList = new ArrayList<Field>();
        Column columnAnno = null;
        for (Field field : fields) {
            columnAnno = field.getAnnotation(Column.class);
            if (columnAnno != null) {
                fieldList.add(field);
            }
        }
        fieldList.sort(new Comparator<Field>() {
            @Override
            public int compare(Field o1, Field o2) {
                return o1.getAnnotation(Column.class).sortNo() - o2.getAnnotation(Column.class).sortNo();
            }
        });
        return fieldList;
    }

    /**
     * 获取excel工作组对象
     */
    private void getWorkbook() {
        this.workbook = new HSSFWorkbook();
    }

    /**
     * 生成每一个sheet的表格头
     * @param fieldList 字段集合
     */
    private void createSheetHeaderRow(List<Field> fieldList) {
        HSSFSheet sheet = workbook.getSheet(sheetName);
        HSSFRow rowHeadNo = sheet.createRow(0);
        int cellHeadNo = 0;
        Column column;
        for (Field field : fieldList) {
            sheet.setColumnWidth(cellHeadNo, 4000);
            HSSFCell cell = rowHeadNo.createCell(cellHeadNo);
            column = field.getAnnotation(Column.class);
            cell.setCellValue(column.name());
            cell.setCellStyle(this.setCellStyle( PART_HEAD));
            cellHeadNo++;
        }
    }

    /**
     * 根据一个对象生成一行导出数据
     * @param rowNo sheet中的行号
     * @param entityObj 需要导出的对象数据
     * @param fieldList 字段列表
     */
    private void createSheetBodyRow(int rowNo, Object entityObj, List<Field> fieldList,HSSFPatriarch patriarch) throws ExportExcelException{
        HSSFSheet sheet = workbook.getSheet(sheetName);
        HSSFRow row = sheet.createRow(rowNo);
        int cellBodyNo = 0;
        short index = 0;
        short imgHeight = 0;
        setCellValue(rowNo, entityObj, fieldList, patriarch, row, cellBodyNo, index, imgHeight);
    }

    private void setCellValue(int rowNo, Object entityObj, List<Field> fieldList, HSSFPatriarch patriarch, HSSFRow row, int cellBodyNo, short index, short imgHeight) throws ExportExcelException {
        HSSFCell cell;
        Column column;
        for (Field field : fieldList) {
            cell = row.createCell(cellBodyNo);
            column = field.getAnnotation(Column.class);
            field.setAccessible(true);
            try {
                imgHeight = formatCellValue(rowNo, entityObj, patriarch, row, index, imgHeight, cell, column, field);
            } catch (Exception e) {
                e.printStackTrace();
                throw new ExportExcelException(e);
            }
            cellBodyNo++;
            index ++;
        }
    }

    private short formatCellValue(int rowNo, Object entityObj, HSSFPatriarch patriarch, HSSFRow row, short index, short imgHeight, HSSFCell cell, Column column, Field field) throws IllegalAccessException, InvocationTargetException, NoSuchMethodException, IOException {
        String simpleName = field.getType().getSimpleName();
        this.imgType = column.imgType();
        if (column.formatter() != null && !column.formatter().equals("")) {
            cell.setCellValue(field.get(entityObj) != null ? entityObj.getClass().getMethod(column.formatter(), field.getType()).invoke(entityObj, field.get(entityObj)).toString() : "");
        } else {
            if(simpleName.equals("Integer")||simpleName.equals("int")){
                cell.setCellValue(field.get(entityObj)!=null?Integer.valueOf(field.get(entityObj).toString()):0);
            }else if(simpleName.equals("Long")||simpleName.equals("long")){
                cell.setCellValue(field.get(entityObj)!=null?Long.valueOf(field.get(entityObj).toString()):0L);
            }else if(simpleName.equals("Byte")||simpleName.equals("byte")){
                cell.setCellValue(field.get(entityObj)!=null?Byte.valueOf(field.get(entityObj).toString()):0);
            }else if(simpleName.equals("Character")||simpleName.equals("char")){
                cell.setCellValue(field.get(entityObj)!=null?Character.valueOf(field.get(entityObj).toString().charAt(0)):Character.MIN_VALUE);
            }else if(simpleName.equals("Short")||simpleName.equals("short")){
                cell.setCellValue(field.get(entityObj)!=null?Short.valueOf(field.get(entityObj).toString()):0);
            }else if(simpleName.equals("byte[]")){
                if(field.get(entityObj)==null){
                    cell.setCellValue("");
                }else{
                    File file = FileUtil.byte2file((byte[])field.get(entityObj),this.imgType);
                    if(!file.exists()){
                        file.createNewFile();
                    }
                    imgHeight = exportImage(rowNo, patriarch, row, index, imgHeight, file);
                }
            }else if(simpleName.equals("File")){
                if(field.get(entityObj)==null){
                    cell.setCellValue("");
                }else{
                    imgHeight = exportImage(rowNo, patriarch, row, index, imgHeight, (File)field.get(entityObj));
                }
            }else{
                cell.setCellValue(field.get(entityObj)!=null?field.get(entityObj).toString():"");
            }
        }
        cell.setCellStyle(this.setCellStyle( ExcelUtil.PART_BODY));
        return imgHeight;
    }

    private short exportImage(int rowNo, HSSFPatriarch patriarch, HSSFRow row, short index, short imgHeight, File file) throws IOException {
        ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
        BufferedImage bufferImg = ImageIO.read(file);
        if((short)(bufferImg.getHeight()*10)>imgHeight){
            imgHeight = Integer.valueOf(bufferImg.getHeight()*10).shortValue();
        }
        row.setHeight(imgHeight);
        HSSFClientAnchor anchor =new HSSFClientAnchor(0,0,1020,250,index,rowNo,index,rowNo);//dx2最大值  1023,dy2最大值255
        if(this.imgType!=null&&!this.imgType.equals("")&&this.imgType.equals("jpg")){
            ImageIO.write(bufferImg, "jpg", byteArrayOut);
            patriarch.createPicture(anchor,this.workbook.addPicture(byteArrayOut.toByteArray(), HSSFWorkbook.PICTURE_TYPE_JPEG));
        }else if(this.imgType!=null&&!this.imgType.equals("")&&this.imgType.equals("png")){
            ImageIO.write(bufferImg, "png", byteArrayOut);
            patriarch.createPicture(anchor,this.workbook.addPicture(byteArrayOut.toByteArray(), HSSFWorkbook.PICTURE_TYPE_PNG));
        }
        byteArrayOut.close();
        file.delete();
        return imgHeight;
    }

    /**
     * 设置单元格的样式
     * @param part 样式类别
     * @return HSSFCellStyle 单元格的样式对象
     */
    private HSSFCellStyle setCellStyle( String part) {
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        HSSFFont font = workbook.createFont();
        if (ExcelUtil.PART_HEAD.equals(part)) {
            font.setColor(Font.COLOR_RED);
        }
        if (ExcelUtil.PART_BODY.equals(part)) {
            font.setColor(Font.COLOR_NORMAL);
        }
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        cellStyle.setFont(font);
        cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);//垂直居中
        return cellStyle;
    }

    /**
     * 生成excel文件
     */
    private void createExcelFile(String exportFilePath) throws ExportExcelException{
        String fileName = this.fileName + new SimpleDateFormat("yyyyMMddHHmmss").format(new Date()) + ".xls";
        OutputStream out = null;
        try {
            out = new FileOutputStream(new File(exportFilePath + fileName));
            workbook.write(out);
        } catch (Exception e) {
            e.printStackTrace();
            throw new ExportExcelException("excel导出异常");
        } finally {
            try {
                out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}