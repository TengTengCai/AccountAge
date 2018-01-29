package com.tongguan.excel1_2;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PropertyTemplate;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class AccountAge {
    private String projectPath = "";
    private String inputFilePath = "E:\\IDEAWorkSpace\\excel1_2\\src\\com\\tongguan\\inputFile\\三年其他应收账款.xlsx"; //输入处理文件的路径地址
    private String outputFilePath = "E:\\IDEAWorkSpace\\excel1_2\\src\\com\\tongguan\\outputFile\\result.xlsx";//输出结果文件的路径地址
    private FileController fileController = new FileController();   //文件控制
    private XSSFWorkbook sWorkBook;     //数据资源工作薄
    private XSSFWorkbook rWorkBook;     //结果输出工作薄
    private XSSFSheet rSheet;           //结果分页
    private int sheetNumbers = 0;       //数据资源工作簿分页数量
    private int coordinateRuler = 0;    //标尺，不同的年份表格标尺不一样
    private int sheetWidth;             //表格宽度
    private List<Cell> sList;           //客户list列表
    private Map<String, Double> sheetMap = new HashMap<>();
    private CellStyle cellStyle;        //单元格样式
    private boolean isReceivable = true;//应收为true 应付为false
    private String subjectName = "";    //项目名称，通过项目名称来判断应收，还是应付

    /**
     * 无参构造方法
     */
    AccountAge() {

    }

    /**
     * 有参构造方法，初始化参数赋值
     * 传递进来的科目名称进行判断，是付账，还是收账
     *
     * @param subjectName 项目名称，默认文件输入输出路径
     */
    AccountAge(String subjectName) {
        this.subjectName = subjectName;
        this.projectPath = System.getProperty("user.dir");
    }

    /**
     * 有参构造方法，初始化参数赋值
     * 传递进来的科目名称进行判断，是付账，还是收账
     *
     * @param inputFilePath  输入文件的路径
     * @param outputFilePath 输出文件袋路径
     * @param subjectName    项目名称
     */
    AccountAge(String inputFilePath, String outputFilePath, String subjectName) {
        this.inputFilePath = inputFilePath;
        this.outputFilePath = outputFilePath;
        this.subjectName = subjectName;
        this.projectPath = System.getProperty("user.dir");
    }

    /**
     * 初始化方法，流程控制
     */
    public void doInit() {
        System.out.println(projectPath);
        String[] strArray = subjectName.split("付");     //通过split来判断是否有付字
        if (strArray.length >= 2) {           //有付字，字符串被分割成一个数组，长度大于等于2，没有数组长度则为1
            isReceivable = false;           //应收应付判断参数，true为应收，false为应付
        }
        sWorkBook = fileController.getFile(inputFilePath);//获取数据源工作薄
        sheetNumbers = sWorkBook.getNumberOfSheets();       //数据资源工作簿的分页多少
        coordinateRuler = sheetNumbers*2+2;
        String mouldFilePath = "";                          //模版文件路径
        switch (sheetNumbers) {                             //判断是几年使用什么模版的文件
            case 3:
                mouldFilePath = projectPath + "\\src\\com\\tongguan\\mouldFile\\mould3.xlsx";
                break;
            case 4:
                mouldFilePath = projectPath + "\\src\\com\\tongguan\\mouldFile\\mould4.xlsx";
                break;
            case 5:
                mouldFilePath = projectPath + "\\src\\com\\tongguan\\mouldFile\\mould5.xlsx";
                break;
            default:
                break;
        }

        rWorkBook = fileController.copyFile(mouldFilePath, outputFilePath); //复制模版文件
        cellStyle = rWorkBook.createCellStyle();                            //设置一个全局的数据格式
        XSSFDataFormat format = rWorkBook.createDataFormat();
        cellStyle.setDataFormat(format.getFormat("#,##0.00"));              //小数点后两位，千分格
        sList = new ArrayList<>();                                          //新建一个list对象，存储客户名称
        initCustomerList();             //初始化客户列表，和年份，项目名称

        getDataToMap();                 //获取数据
        putDataFromMap();               //输入数据

        setFormula();                   //设置函数
        //设置边框
        PropertyTemplate propertyTemplate = new PropertyTemplate();
        propertyTemplate.drawBorders(new CellRangeAddress(4, rSheet.getLastRowNum() - 2, 0, sheetWidth),
                BorderStyle.THIN, BorderExtent.ALL);
        propertyTemplate.applyBorders(rSheet);
//
        //设置自动宽度
        for (int i = 0; i < 18; i++) {
            rSheet.autoSizeColumn(i);
        }

        //总数和的校验
        Row row = rSheet.getRow(rSheet.getLastRowNum());
        Cell cell = row.createCell(0);
        cell.setCellValue("审计说明：");

        fileController.createFile(outputFilePath, rWorkBook);//另存为，输出的文件
    }

    /**
     * 初始化客户列表
     */
    private void initCustomerList() {
        //获取客户列表
        XSSFSheet sSheet = sWorkBook.getSheetAt(sheetNumbers - 1);
        sList = getColumnWithCol(sSheet, 0);
        sList.remove(0);
        for (Cell cell : sList) {
            if (cell != null) {
                cell.setCellType(CellType.STRING);
//                System.out.println(cell.getStringCellValue());
            } else {
                sList.remove(null);
            }
        }

        //初始化客户列表
        rSheet = rWorkBook.getSheetAt(0);
        for (int i = 0; i < sList.size(); i++) {
            Row row = rSheet.createRow(i + 6);
            Cell cell = row.createCell(0);
            cell.setCellType(CellType.STRING);
            cell.setCellValue(sList.get(i).getStringCellValue());
        }

        //初始化年份
        Row row = rSheet.getRow(4);
        int temp = 1;
        for (int i = 0; i < sheetNumbers; i++) {
            if (i == 0) {
                for (int j = 0; j < 3; j++) {
                    row.getCell(temp).setCellValue(sWorkBook.getSheetName(i) + "年");
                    temp++;
                }
            } else if (i == sheetNumbers - 1) {
                for (int j = 0; j < 3; j++) {
                    row.getCell(temp).setCellValue(sWorkBook.getSheetName(i) + "年");
                    temp++;
                }
            } else {
                for (int j = 0; j < 2; j++) {
                    row.getCell(temp).setCellValue(sWorkBook.getSheetName(i) + "年");
                    temp++;
                }
            }
        }
        //初始化项目名称
        if (!subjectName.equals("")) {
            rSheet.getRow(1).getCell(1).setCellValue(subjectName);
        }
    }

    /**
     * 获取数据存到HashMap中
     */
    private void getDataToMap() {
        int sheetId = 0;
        for (Sheet sheet : sWorkBook) {
            for (int i = 1; i < sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                String customer = row.getCell(0).getStringCellValue();
                for (int j = 3; j < row.getLastCellNum(); j++) {
                    Cell cell = row.getCell(j);
                    double value = cell.getNumericCellValue();
                    if(cell.getCellTypeEnum() == CellType.NUMERIC){
                        if (sheetId == 0){
                            if (j == 3){
                                sheetMap.put(sheet.getSheetName()+"年期初余额"+customer,value);
//                                System.out.println();
                            }
                        }else if (sheetId == sWorkBook.getNumberOfSheets()-1){
                            if (j == 6){
                                sheetMap.put(sheet.getSheetName()+"年期末余额"+customer,value);
//                                System.out.println();
                            }
                        }
                        if (j == 4){
                            sheetMap.put(sheet.getSheetName()+"年借方金额"+customer,value);
//                                System.out.println();
                        }else if (j == 5){
                            sheetMap.put(sheet.getSheetName()+"年贷方金额"+customer,value);
//                                System.out.println();
                        }

                    }
                }
            }
            sheetId++;
        }

    }

    /**
     * 把HashMap的数据输入数据到表格中
     */
    private void putDataFromMap(){
        Row yearRow = rSheet.getRow(4);
        Row titleRow = rSheet.getRow(5);
        for (int i = 6 ; i <= rSheet.getLastRowNum();i++){
            Row row = rSheet.getRow(i);
            String customer = row.getCell(0).getStringCellValue();
            for (int j = 1; j<= coordinateRuler;j++){
                String year = yearRow.getCell(j).getStringCellValue();
                String title = titleRow.getCell(j).getStringCellValue();
//                System.out.println(year+title+customer);
                Cell cell = row.createCell(j);
                Double value = sheetMap.get(year+title+customer);
                if (value != null){
                    cell.setCellValue(value);
//                    System.out.println(year+title+customer+"-----"+ value);
                }else {
                    cell.setCellValue(0.00);
                }
                cell.setCellStyle(cellStyle);
            }
        }
    }


    /**
     * 将函数设置到对应的单元格中
     */
    private void setFormula() {
        sheetWidth = rSheet.getRow(5).getLastCellNum() - 1;
        if (isReceivable) {
            //应收账款的函数
            switch (sheetNumbers) {
                case 3:
                    setThreeYearFormulaWithIncome();
                    break;
                case 4:
                    setFourYearFormulaWithIncome();
                    break;
                case 5:
                    setFiveYearFormulaWithIncome();
                    break;
                default:
                    break;
            }
        } else {
            //应付账款的函数
            switch (sheetNumbers) {
                case 3:
                    setThreeYearFormulaWithPay();
                    break;
                case 4:
                    setFourYearFormulaWithPay();
                    break;
                case 5:
                    setFiveYearFormulaWithPay();
                    break;
                default:
                    break;
            }
        }

        //所有项目的和
        Row row = rSheet.getRow(5);
        int lastRowNum = rSheet.getLastRowNum();
        Row sumRow = rSheet.createRow(rSheet.getLastRowNum() + 1);
        char c = 66;
        for (int i = 1; i < row.getLastCellNum(); i++) {
            Cell cell = sumRow.createCell(i);
            cell.setCellStyle(cellStyle);
            String str = String.valueOf(c);
            cell.setCellFormula("SUM(" + str + "7:" + str + (lastRowNum + 1) + ")");
            c++;
        }


        row = rSheet.createRow(lastRowNum + 3);
        Cell cell = row.createCell(1);
        int index = lastRowNum + 2;

        //审计说明的函数计算
        cell.setCellStyle(cellStyle);
        if (isReceivable) {
            //应收账款的函数
            switch (sheetNumbers) {
                case 3:
                    cell.setCellFormula("B" + index + "+C" + index + "-D" + index + "+E" + index + "-F" + index + "+G" + index + "-H" + index + "-I" + index);
                    break;
                case 4:
                    cell.setCellFormula("B" + index + "+C" + index + "-D" + index + "+E" + index + "-F" + index + "+G" + index + "-H" + index + "+I" + index + "-J" + index + "-K" + index);
                    break;
                case 5:
                    cell.setCellFormula("B" + index + "+C" + index + "-D" + index + "+E" + index + "-F" + index + "+G" + index + "-H" + index + "+I" + index + "-J" + index + "+K" + index + "-L" + index + "-M" + index);
                    break;
                default:
                    break;
            }
        } else {
            //应付账款的函数
            switch (sheetNumbers) {
                case 3:
                    cell.setCellFormula("B" + index + "-C" + index + "+D" + index + "-E" + index + "+F" + index + "-G" + index + "+H" + index + "-I" + index);
                    break;
                case 4:
                    cell.setCellFormula("B" + index + "-C" + index + "+D" + index + "-E" + index + "+F" + index + "-G" + index + "+H" + index + "-I" + index + "+J" + index + "-K" + index);
                    break;
                case 5:
                    cell.setCellFormula("B" + index + "-C" + index + "+D" + index + "-E" + index + "+F" + index + "-G" + index + "+H" + index + "-I" + index + "+J" + index + "-K" + index + "+L" + index + "-M" + index);
                    break;
                default:
                    break;
            }
        }
    }

    /**
     * 五年的收账函数
     */
    private void setFiveYearFormulaWithIncome() {
        coordinateRuler = 12;
//        Row row = rSheet.getRow(6);
        for (int i = 6; i <= rSheet.getLastRowNum(); i++) {
            Row row = rSheet.getRow(i);
            Cell cell = row.createCell(coordinateRuler);
            cell.setCellStyle(cellStyle);
            int index = i + 1;
            cell.setCellFormula("B" + index + "+C" + index + "-D" + index + "+E" + index + "-F" + index + "+G" + index + "-H" + index + "+I" + index + "-J" + index + "+K" + index + "-L" + index);
            cell = row.createCell(coordinateRuler + 1);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("IF(L" + index + ">=(M" + index + "+L" + index + "-K" + index + "),M" + index + ",K" + index + ")");
            cell = row.createCell(coordinateRuler + 2);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("IF(L" + index + "+J" + index + ">=(M" + index + "+L" + index + "-K" + index + "+J" + index + "-I" + index + "),M" + index + "-N" + index + ",I" + index + ")");
            cell = row.createCell(coordinateRuler + 3);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("IF(L" + index + "+J" + index + "+H" + index + ">=(M" + index + "+L" + index + "-K" + index + "+J" + index + "-I" + index + "+H" + index + "-G" + index + "),M" + index + "-O" + index + "-N" + index + ",G" + index + ")");
            cell = row.createCell(coordinateRuler + 4);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("IF(L" + index + "+J" + index + "+H" + index + "+F" + index + ">=(M" + index + "+L" + index + "-K" + index + "+J" + index + "-I" + index + "+H" + index + "-G" + index + "+F" + index + "-E" + index + "),M" + index + "-P" + index + "-O" + index + "-N" + index + ",E" + index + ")");
            cell = row.createCell(coordinateRuler + 5);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("IF(L" + index + "+J" + index + "+H" + index + "+F" + index + "+D" + index + ">=(M" + index + "+L" + index + "-K" + index + "+J" + index + "-I" + index + "+H" + index + "-G" + index + "+F" + index + "-E" + index + "+D" + index + "-C" + index + "),M" + index + "-P" + index + "-O" + index + "-N" + index + "-Q" + index + ",C" + index + ")");
            cell = row.createCell(coordinateRuler + 6);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("M" + index + "-N" + index + "-O" + index + "-P" + index + "-Q" + index + "-R" + index);
        }
    }

    /**
     * 五年的付账函数
     */
    private void setFiveYearFormulaWithPay() {
        coordinateRuler = 12;
//        Row row = rSheet.getRow(6);
        for (int i = 6; i <= rSheet.getLastRowNum(); i++) {
            Row row = rSheet.getRow(i);
            Cell cell = row.createCell(coordinateRuler);
            cell.setCellStyle(cellStyle);
            int index = i + 1;
            cell.setCellFormula("B" + index + "-C" + index + "+D" + index + "-E" + index + "+F" + index + "-G" + index + "+H" + index + "-I" + index + "+J" + index + "-K" + index + "+L" + index);
            cell = row.createCell(coordinateRuler + 1);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("IF(K" + index + ">=(M" + index + "-L" + index + "+K" + index + "),M" + index + ",L" + index + ")");
            cell = row.createCell(coordinateRuler + 2);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("IF(K" + index + "+I" + index + ">=(M" + index + "-L" + index + "+K" + index + "-J" + index + "+I" + index + "),M" + index + "-N" + index + ",J" + index + ")");
            cell = row.createCell(coordinateRuler + 3);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("IF(K" + index + "+I" + index + "+G" + index + ">=(M" + index + "-L" + index + "+K" + index + "-J" + index + "+I" + index + "-H" + index + "+G" + index + "),M" + index + "-O" + index + "-N" + index + ",H" + index + ")");
            cell = row.createCell(coordinateRuler + 4);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("IF(K" + index + "+I" + index + "+G" + index + "+E" + index + ">=(M" + index + "-L" + index + "+K" + index + "-J" + index + "+I" + index + "-H" + index + "+G" + index + "-F" + index + "+E" + index + "),M" + index + "-P" + index + "-O" + index + "-N" + index + ",F" + index + ")");
            cell = row.createCell(coordinateRuler + 5);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("IF(K" + index + "+I" + index + "+G" + index + "+E" + index + "+C" + index + ">=(M" + index + "-L" + index + "+K" + index + "-J" + index + "+I" + index + "-H" + index + "+G" + index + "-F" + index + "+E" + index + "-D" + index + "+C" + index + "),M" + index + "-P" + index + "-O" + index + "-N" + index + "-Q" + index + ",D" + index + ")");
            cell = row.createCell(coordinateRuler + 6);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("M" + index + "-N" + index + "-O" + index + "-P" + index + "-Q" + index + "-R" + index);
        }
    }

    /**
     * 四年的收账函数
     */
    private void setFourYearFormulaWithIncome() {
        coordinateRuler = 10;
        for (int i = 6; i <= rSheet.getLastRowNum(); i++) {
            Row row = rSheet.getRow(i);
            Cell cell = row.createCell(coordinateRuler);
            cell.setCellStyle(cellStyle);
            int index = i + 1;
            cell.setCellFormula("B" + index + "+C" + index + "-D" + index + "+E" + index + "-F" + index + "+G" + index + "-H" + index + "+I" + index + "-J" + index);
            cell = row.createCell(coordinateRuler + 1);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("IF(J" + index + ">=(K" + index + "+J" + index + "-I" + index + "),K" + index + ",I" + index + ")");
            cell = row.createCell(coordinateRuler + 2);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("IF(J" + index + "+H" + index + ">=(K" + index + "+J" + index + "-I" + index + "+H" + index + "-G" + index + "),K" + index + "-L" + index + ",G" + index + ")");
            cell = row.createCell(coordinateRuler + 3);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("IF(J" + index + "+H" + index + "+F" + index + ">=(K" + index + "+J" + index + "-I" + index + "+H" + index + "-G" + index + "+F" + index + "-E" + index + "),K" + index + "-M" + index + "-L" + index + ",E" + index + ")");
            cell = row.createCell(coordinateRuler + 4);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("IF(J" + index + "+H" + index + "+F" + index + "+D" + index + ">=(K" + index + "+J" + index + "-I" + index + "+H" + index + "-G" + index + "+F" + index + "-E" + index + "+D" + index + "-C" + index + "),K" + index + "-N" + index + "-M" + index + "-L" + index + ",C" + index + ")");
            cell = row.createCell(coordinateRuler + 5);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("K" + index + "-L" + index + "-M" + index + "-N" + index + "-O" + index);
        }
    }

    /**
     * 四年的付账函数
     */
    private void setFourYearFormulaWithPay() {
        coordinateRuler = 10;
        for (int i = 6; i <= rSheet.getLastRowNum(); i++) {
            Row row = rSheet.getRow(i);
            Cell cell = row.createCell(coordinateRuler);
            cell.setCellStyle(cellStyle);
            int index = i + 1;
            cell.setCellFormula("B" + index + "-C" + index + "+D" + index + "-E" + index + "+F" + index + "-G" + index + "+H" + index + "-I" + index + "+J" + index);
            cell = row.createCell(coordinateRuler + 1);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("IF(I" + index + ">=(K" + index + "-J" + index + "+I" + index + "),K" + index + ",J" + index + ")");
            cell = row.createCell(coordinateRuler + 2);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("IF(I" + index + "+G" + index + ">=(K" + index + "-J" + index + "+I" + index + "-H" + index + "+G" + index + "),K" + index + "-L" + index + ",H" + index + ")");
            cell = row.createCell(coordinateRuler + 3);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("IF(I" + index + "+G" + index + "+E" + index + ">=(K" + index + "-J" + index + "+I" + index + "-H" + index + "+G" + index + "-F" + index + "+E" + index + "),K" + index + "-M" + index + "-L" + index + ",F" + index + ")");
            cell = row.createCell(coordinateRuler + 4);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("IF(I" + index + "+G" + index + "+E" + index + "+C" + index + ">=(K" + index + "-J" + index + "+I" + index + "-H" + index + "+G" + index + "-F" + index + "+E" + index + "-D" + index + "+C" + index + "),K" + index + "-N" + index + "-M" + index + "-L" + index + ",D" + index + ")");
            cell = row.createCell(coordinateRuler + 5);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("K" + index + "-L" + index + "-M" + index + "-N" + index + "-O" + index);
        }
    }

    /**
     * 三年的收账函数
     */
    private void setThreeYearFormulaWithIncome() {
        coordinateRuler = 8;
        for (int i = 6; i <= rSheet.getLastRowNum(); i++) {
            Row row = rSheet.getRow(i);
            Cell cell = row.createCell(coordinateRuler);
            cell.setCellStyle(cellStyle);
            int index = i + 1;
            cell.setCellFormula("B" + index + "+C" + index + "-D" + index + "+E" + index + "-F" + index + "+G" + index + "-H" + index);
            cell = row.createCell(coordinateRuler + 1);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("IF(H" + index + ">=(I" + index + "+H" + index + "-G" + index + "),I" + index + ",G" + index + ")");
            cell = row.createCell(coordinateRuler + 2);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("IF(H" + index + "+F" + index + ">=(I" + index + "+H" + index + "-G" + index + "+F" + index + "-E" + index + "),I" + index + "-J" + index + ",E" + index + ")");
            cell = row.createCell(coordinateRuler + 3);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("IF(H" + index + "+F" + index + "+D" + index + ">=(I" + index + "+H" + index + "-G" + index + "+F" + index + "-E" + index + "+D" + index + "-C" + index + "),I" + index + "-K" + index + "-J" + index + ",C" + index + ")");
            cell = row.createCell(coordinateRuler + 4);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("I" + index + "-J" + index + "-K" + index + "-L" + index);
        }
    }

    /**
     * 三年的付账函数
     */
    private void setThreeYearFormulaWithPay() {
        coordinateRuler = 8;
        for (int i = 6; i <= rSheet.getLastRowNum(); i++) {
            Row row = rSheet.getRow(i);
            Cell cell = row.createCell(coordinateRuler);
            cell.setCellStyle(cellStyle);
            int index = i + 1;
            cell.setCellFormula("B" + index + "-C" + index + "+D" + index + "-E" + index + "+F" + index + "-G" + index + "+H" + index);
            cell = row.createCell(coordinateRuler + 1);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("IF(G" + index + ">=(I" + index + "-H" + index + "+G" + index + "),I" + index + ",H" + index + ")");
            cell = row.createCell(coordinateRuler + 2);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("IF(G" + index + "+E" + index + ">=(I" + index + "-H" + index + "+G" + index + "-F" + index + "+E" + index + "),I" + index + "-J" + index + ",F" + index + ")");
            cell = row.createCell(coordinateRuler + 3);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("IF(G" + index + "+E" + index + "+C" + index + ">=(I" + index + "-H" + index + "+G" + index + "-F" + index + "+E" + index + "-D" + index + "+C" + index + "),I" + index + "-K" + index + "-J" + index + ",D" + index + ")");
            cell = row.createCell(coordinateRuler + 4);
            cell.setCellStyle(cellStyle);
            cell.setCellFormula("I" + index + "-J" + index + "-K" + index + "-L" + index);
        }
    }

    //一系列变量的set和get函数
    public void setSubjectName(String subjectName) {
        this.subjectName = subjectName;
    }

    public String getSubjectName() {
        return subjectName;
    }

    public String getInputFilePath() {
        return inputFilePath;
    }

    public void setInputFilePath(String inputFilePath) {
        this.inputFilePath = inputFilePath;
    }

    public String getOutputFilePath() {
        return outputFilePath;
    }

    public void setOutputFilePath(String outputFilePath) {
        this.outputFilePath = outputFilePath;
    }

    private List<Cell> getColumnWithCol(Sheet sheet, int col) {
        List<Cell> sheetCol = new ArrayList<>();
        for (Row row : sheet) {
            Cell cell = row.getCell(col);
            if (cell != null) {
                CellType cellType = cell.getCellTypeEnum();
                switch (cellType) {
                    case STRING:
                        String str = cell.getStringCellValue();
                        if (!str.isEmpty()) {
                            sheetCol.add(cell);
                        }
                        break;
                    case NUMERIC:
                        sheetCol.add(cell);
                        break;
                    default:
                        break;
                }
            }
        }
        return sheetCol;
    }
}
