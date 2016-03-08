package exportXLS.util;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Date;
import java.util.List;

import exportXLS.exception.ExportExcelException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Font;

import exportXLS.annotation.Column;
import exportXLS.annotation.Xls;


public class ExcelUtil {
    private String sheetName = "";
    private final static String PART_HEAD = "head";
    private final static String PART_BODY = "body";
    private HSSFWorkbook workbook;
    private final static Integer MAX_PAGE_NUM = 43000;

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
        String sheetName = excelClass.name();
        int size = entityList.size();
        try {
            //创建工作页的行，根据list中的对象创建数据行，从滴1行开始
            if (size > 0) {
                pageCount = (size / MAX_PAGE_NUM) + 1;
                for (int i = 0; i < pageCount; i++) {
                    this.sheetName = sheetName + (i + 1);
                    this.workbook.createSheet(this.sheetName);
                    this.createSheetHeaderRow(fieldList);
                    for (int rowNo = 1, listNo = i * MAX_PAGE_NUM; listNo < (((i + 1) == pageCount) ? (i * MAX_PAGE_NUM) + (size % MAX_PAGE_NUM) : (i + 1) * MAX_PAGE_NUM); rowNo++, listNo++) {
                        this.createSheetBodyRow(rowNo, entityList.get(listNo), fieldList);
                    }
                }
            }
            createExcelFile();
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
    private void createSheetBodyRow(int rowNo, Object entityObj, List<Field> fieldList) throws ExportExcelException{
        HSSFSheet sheet = workbook.getSheet(sheetName);
        HSSFRow rowBodyNo = sheet.createRow(rowNo);
        int cellBodyNo = 0;
        Column column;
        HSSFCell cell;
        for (Field field : fieldList) {
            cell = rowBodyNo.createCell(cellBodyNo);
            column = field.getAnnotation(Column.class);
            field.setAccessible(true);
            try {
                String cellValue = "";
                if (column.formatter() != null && !column.formatter().equals("")) {
                    cellValue = field.get(entityObj) != null ? entityObj.getClass().getMethod(column.formatter(), field.getType()).invoke(entityObj, field.get(entityObj)).toString() : "";
                } else {
                    cellValue = field.get(entityObj) != null ? field.get(entityObj).toString() : "";
                }
                cell.setCellValue(cellValue);
                cell.setCellStyle(this.setCellStyle( ExcelUtil.PART_BODY));
            } catch (Exception e) {
                e.printStackTrace();
                throw new ExportExcelException(e);
            }
            cellBodyNo++;
        }
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
        return cellStyle;
    }

    /**
     * 生成excel文件
     */
    private void createExcelFile() throws ExportExcelException{
        String fileName = sheetName + new SimpleDateFormat("yyyyMMddHHmmss").format(new Date()) + ".xls";
        OutputStream out = null;
        try {
            out = new FileOutputStream(new File("d:/" + fileName));
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

