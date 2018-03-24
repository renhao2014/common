package com.template.util.file.excel;


import com.alibaba.fastjson.JSON;
import com.template.util.model.Template;
import com.template.util.reflection.BeanHelper;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.beans.IntrospectionException;
import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.util.*;

/**
 * 陈仁浩
 * 重新封装了Excel工具，支持Xssf和Hssf
 */
public class ExcelHelper {

    // 默认高度
    private static short DEFAULT_ROW_HEIGHT = 400;
    // 默认宽度
    private static int DEFAULT_CELL_WIDTH = 3000;



    /****************************导出部分**********************************
     *
    *****************************导出部分*********************************/


    /**
     * 获取默认数据单元格样式
     */
    protected static CellStyle getDefaultDataStyle(Workbook workbook){
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setBorderBottom(CellStyle.BORDER_THIN); // 下边框
        style.setBorderLeft(CellStyle.BORDER_THIN);// 左边框
        style.setBorderTop(CellStyle.BORDER_THIN);// 上边框
        style.setBorderRight(CellStyle.BORDER_THIN);// 右边框
        return style;
    }

    /**
     * 获取默认标题单元格样式
     * @return
     */
    protected static CellStyle getDefaultTitleStyle(Workbook workbook){
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setBorderBottom(CellStyle.BORDER_THIN); // 下边框
        style.setBorderLeft(CellStyle.BORDER_THIN);// 左边框
        style.setBorderTop(CellStyle.BORDER_THIN);// 上边框
        style.setBorderRight(CellStyle.BORDER_THIN);// 右边框
        return style;
    }

    /**
     * Row中创建新Cell,并按指定样式写值
     * @param row
     * @param index
     * @param value
     * @param cellStyle 有默认样式，可不传
     * @return
     */
    protected static Cell setNewCell(Row row,Integer index,String value,CellStyle cellStyle){
        Cell newCell = row.createCell(index);
        newCell.setCellType(Cell.CELL_TYPE_STRING);
        newCell.setCellValue(value);
        if (cellStyle==null){
            newCell.setCellStyle(getDefaultDataStyle(row.getSheet().getWorkbook()));
        }else{
            newCell.setCellStyle(cellStyle);
        }
        return newCell;
    }


    /**
     * 追加基本行
     * @param sheet
     * @param isFirstRow
     * @param segs
     * @param cellStyle 有默认样式，可不传
     */
    public static void appendStringsRow(Sheet sheet,Boolean isFirstRow, String[] segs,CellStyle cellStyle){
        Integer lastRow = sheet.getLastRowNum();
        if (!isFirstRow) {
            lastRow = lastRow + 1;
        }
        Row newRow = sheet.createRow(lastRow);
        int index = 0;
        for (String seg : segs) {
            setNewCell(newRow,index,seg,cellStyle);
            index++;
        }
    }


    /**
     * 追加map指定属性s的值行
     * @param sheet
     * @param isFirstRow
     * @param map
     * @param attrs
     * @param cellStyle 有默认样式，可不传
     */
    public static void appendMapRow(Sheet sheet, Boolean isFirstRow, Map<String,Object> map,String[] attrs, CellStyle cellStyle){
        String[] segs= new String[attrs.length];
        for (int i=0;i<attrs.length;i++) {
            segs[i]= BeanHelper.convertObjToString(map.get(attrs[i]));;
        }
        appendStringsRow(sheet,isFirstRow,segs,cellStyle);
    }


    /**
     * 追加obj指定属性s的值行
     * @param sheet
     * @param isFirstRow
     * @param obj
     * @param attrs
     * @param cellStyle 有默认样式，可不传
     * @throws IllegalAccessException
     * @throws IntrospectionException
     * @throws InvocationTargetException
     */
    public static void appendObjectRow(Sheet sheet, Boolean isFirstRow, Object obj, String[] attrs, CellStyle cellStyle) throws IllegalAccessException, IntrospectionException, InvocationTargetException {
        Map<String, Object> map = BeanHelper.convertBeanToMap(obj);
        appendMapRow(sheet, isFirstRow, map, attrs, cellStyle);
    }

    /**
     * 内部自建workbook，不支持自定义表头
     * 默认表数据导出(传入list<obj/map>数据+非已有sheet)
     * 默认表数据导出(xlsx格式)
     * 如果修改请继承类然后修改
     */
    public static Workbook defaultExport(String sheetTitle, String[] titles, String[] attrs,
                              List<Object> objs) throws IllegalAccessException, IntrospectionException, InvocationTargetException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(sheetTitle);
        appendStringsRow(sheet,true,titles,getDefaultTitleStyle(workbook));

        for (int i = 0; i < objs.size(); i++) {
            if (objs.get(i) instanceof Map){
                appendMapRow(sheet,false,(Map<String, Object>) objs.get(i),attrs,null);
            }else{
                appendObjectRow(sheet,false,objs.get(i),attrs,null);
            }
        }
        //设置列宽
        for (int i = 0; i < 50; i++) {
            sheet.setColumnWidth(i, 6000);
        }
        return workbook;
    }



    /**
     * 默认表数据导出(传入list<obj/map>数据+已有sheet)
     * 接收外部传入workbook 支持自定义表头
     * 如果修改请继承类然后修改
     * @param sheet
     * @param sheetTitle
     * @param titles
     * @param attrs
     */
    public static Workbook defaultExport(Sheet sheet,boolean isFirstRow,String sheetTitle, String[] titles, String[] attrs,
                              List<Object> objs) throws IllegalAccessException, IntrospectionException, InvocationTargetException {
        appendStringsRow(sheet,isFirstRow,titles,getDefaultTitleStyle(sheet.getWorkbook()));
        for (int i = 0; i < objs.size(); i++) {
            if (objs.get(i) instanceof Map){
                appendMapRow(sheet,false,(Map<String, Object>) objs.get(i),attrs,null);
            }else{
                appendObjectRow(sheet,false,objs.get(i),attrs,null);
            }
        }
        return sheet.getWorkbook();
    }


    /****************************导入部分**********************************
     *
     * 此部分需要适配目标文档，比较定制化比如经常需要考虑合并等情况，
     * 因此仅封装了比较底层的方法
     *
    *****************************导入部分*********************************/


    public static Object parseCellValue(Cell cell){
        if (cell == null) {
            return null;
        }
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                // 字符串中也有符合时间格式的
                return cell.getRichStringCellValue().getString().trim();
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                } else {
                    return cell.getNumericCellValue();
                }
            case Cell.CELL_TYPE_BOOLEAN:
                return cell.getBooleanCellValue();
            case Cell.CELL_TYPE_FORMULA:
                try {
                    return cell.getStringCellValue();
                } catch (IllegalStateException e) {
                    return String.valueOf(cell.getNumericCellValue());
                }
            default:
                return null;
        }
    }


    /**
     * 读取指定行第N个有效数据(跳过空格单元)
     * @param row
     * @param num
     * @return
     */
    public static Object readRowEffctiveCell(Row row, int num){
        int effctiveCellNum=0;
        for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
            Object o = parseCellValue(row.getCell(i));
            if (o!=null){
                effctiveCellNum++;
                if (effctiveCellNum==num){
                    return o;
                }
            }
        }
        return null;
    }


    /**
     * 读取整数类型单元格
     * @param row
     * @param column
     * @return
     */
    public static Integer readIntCell(Sheet sheet,Integer row,Integer column){
        return null;
    }

    /**
     * 读取字符串类型单元格
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    public static String readStringCell(Sheet sheet,Integer row,Integer column){
        return sheet.getRow(row).getCell(column).toString();
    }


    /**
     * 获取行总数
     * @param sheet
     * @return
     */
    public static Integer readRowNum(Sheet sheet){
        return sheet.getLastRowNum();
    }

    /**
     * 读取单行数据
     * @param sheet
     * @param rowIndex
     * @param columnIndexs
     * @return
     */
    public static List<Object> readRow(Sheet sheet,int rowIndex, int[] columnIndexs){
        ArrayList<Object> list = new ArrayList<Object>();
        Row row = sheet.getRow(rowIndex);
        for (int i = 0; i <columnIndexs.length ; i++) {
            list.add(parseCellValue(row.getCell(columnIndexs[i])));
        }
        return list;
    }


    /**
     * 读取单行map数据
     * @param sheet
     * @param rowIndex
     * @param columnIndexs
     * @param attrs
     * @return
     */
    public static Map<String,Object> readMapRow(Sheet sheet,int rowIndex,int[] columnIndexs,String[] attrs){
        HashMap<String, Object> hashMap = new HashMap();
        List<Object> list = readRow(sheet, rowIndex, columnIndexs);
        for (int i = 0; i < list.size(); i++) {
            hashMap.put(attrs[i],list.get(i));
        }
        return hashMap;
    }


    /**
     * 读取单行Object数据
     * @param sheet
     * @param rowIndex
     * @param columnIndexs
     * @param attrs
     * @param clazz
     * @return
     * @throws IllegalAccessException
     * @throws InstantiationException
     */
    public static <T> T readObjectRow(Sheet sheet, int rowIndex, int[] columnIndexs, String[] attrs, Class<T> clazz) throws IllegalAccessException, InstantiationException {
        Map<String, Object> hashMap = readMapRow(sheet, rowIndex, columnIndexs, attrs);
        T t = clazz.newInstance();
        BeanHelper.convertMapToObject(hashMap, t);
        return t;
    }



    /**
     * 从Excel sheet指定行开始加载数据
     * @param sheet
     * @param columnIndexs
     * @param clazz
     * @param attrs
     * @param startRowNum
     * @param <T>
     * @return
     * @throws Exception
     */
    public static <T> List<T> defaultImport(Sheet sheet, int[] columnIndexs,
                                                Class<T> clazz, String[] attrs,Integer startRowNum) throws Exception {
        ArrayList<T> ts = new ArrayList<T>();
        Row row;
        // 遍历rows
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();
        for (int j = firstRowNum; j <= lastRowNum; j++) {
                if (j < startRowNum) {
                continue;
            }
            T t = readObjectRow(sheet, j, columnIndexs, attrs, clazz);
            ts.add(t);
        }
        return ts;
    }


    /**
     * 从Excel文件流中指定行开始加载数据
     * @param inputStream
     * @param columnIndexs
     * @param clazz
     * @param attrs
     * @param startRowNum
     * @param <T>
     * @return
     * @throws Exception
     */
    public static <T> List<T> defaultImport(InputStream inputStream, int[] columnIndexs,
                                            Class<T> clazz, String[] attrs, Integer startRowNum) throws Exception {
        Sheet sheet = WorkbookFactory.create(inputStream).getSheetAt(0);
        List<T> ts = defaultImport(sheet, columnIndexs, clazz, attrs, startRowNum);
        return ts;
    }




    /********************************其他**********************************
     *
    *********************************其他*********************************/
    /**
     * 获取所有sheet名称
     * @param workbook
     * @return
     */
    public static Set<String> getSheetsName(Workbook workbook){
        HashSet<String> sheetsName = new HashSet<String>();
        int numberOfSheets = workbook.getNumberOfSheets();
        for (int i = 0; i < numberOfSheets; i++) {
            sheetsName.add(workbook.getSheetAt(i).getSheetName());
        }
        return sheetsName;
    }

    /**
     * 获取workBook
     * @param inputStream
     * @return
     */
    public static Workbook getWorkBook(InputStream inputStream){
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(inputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return workbook;
    }

    /**
     * 获取sheet
     * @param path
     * @param sheetName
     * @return
     */
    public static Sheet getSheet(String path,String sheetName) {
        Workbook workBook = null;
        Sheet sheet=null;
        try {
            workBook = getWorkBook(new FileInputStream(path));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        if (workBook!=null){
            sheet = workBook.getSheet(sheetName);
        }
        return sheet;
    }

    /****************************测试部分**********************************
     *
    *****************************测试部分**********************************/
    /**
     * 测试导出
     * @throws IllegalAccessException
     * @throws IntrospectionException
     * @throws InvocationTargetException
     * @throws IOException
     */
    private static void testExport() throws IllegalAccessException, IntrospectionException, InvocationTargetException, IOException {
        String sheetTitle="sheet名称";
        String[] titles={"列标题A","列标题B","列标题C"};
        //String[] attrs={"templateString","templateInt","templateDate"};
        String[] attrs={"templateName","templateType","createTime"};

        List<Object> templates = new ArrayList<Object>();
        for (int i = 0; i < 3; i++) {
            Template template = new Template();
            template.setTemplateName("名称");
            template.setTemplateType(222);
            template.setCreateTime(new Date());
            templates.add(template);
        }
        Workbook workbook = defaultExport(sheetTitle, titles, attrs, templates);
        FileOutputStream fileOutputStream = new FileOutputStream(new File("E://test.xlsx"));
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }


    /**
     * 测试导入
     * @throws Exception
     */
    private static void testImport() throws Exception {
        //String[] attrs={"templateString","templateInt","templateDate"};
        String[] attrs={"templateName","templateType","createTime"};

        int[] columnIndexs={0,1,2};
        FileInputStream fileInputStream = new FileInputStream("E://test.xlsx");
        List<Template> templates = defaultImport(fileInputStream, columnIndexs, Template.class, attrs, 1);
        System.out.println(JSON.toJSONString(templates));
    }

    //测试获取workBook所有sheet名称
    private static void testGetSheetsName() throws IOException, InvalidFormatException {
        FileInputStream fileInputStream = new FileInputStream("C:\\Users\\csh\\Desktop\\人大预算\\2017\\2017市本级部门预算\\002市人大办公厅2017年部门预算.xls");
        Workbook workbook = WorkbookFactory.create(fileInputStream);
        Set<String> sheetsName = getSheetsName(workbook);
        System.out.println("读取到的sheetsName:"+sheetsName);
    }

    //测试读取行指定列
    private static void testReadRow(){
        //预算总表
        String filePath="C:\\Users\\csh\\Desktop\\人大预算\\2017\\2017市本级部门预算\\002市人大办公厅2017年部门预算.xls";

        Sheet sheet = getSheet(filePath, "表1部门收支预算总表");
        int[] columnIndexs={2,3};
        List<Object> rowData = ExcelHelper.readRow(sheet, 6, columnIndexs);
        System.out.println(rowData);
    }



    /**
     * 测试-main
     * 注意修改测试方法内的excel路径
     * @param args
     * @throws Exception
     */
    public static void main(String[] args) throws Exception {
        //testExport();
        //testImport();
        //testGetSheetsName();
        testReadRow();
    }



}
