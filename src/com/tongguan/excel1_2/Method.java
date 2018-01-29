package com.tongguan.excel1_2;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Method {

    /**
     * 默认无参构造参数
     */
    Method() {
    }

    /**
     * 获取一列的cell，存储在list
     * @param sheet 分页
     * @param col   列
     * @return  list对象
     */
    public List<Cell> getColumnWithCol(Sheet sheet, int col) {
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

    /**
     * 通过客户名称和标题获取对应的cell
     * @param sheet     分页
     * @param customer  客户
     * @param title     标题
     * @return  对应的单元格
     */
    public Cell selectCellByCustomerAndTitle(Sheet sheet, String customer, String title) {
        Cell cCell = getCellByValue(sheet, customer);
        if (cCell == null) {
            return null;
        }
        Cell tCell = getCellByValue(sheet, title);
        if (tCell == null) {
            return null;
        }
        return selectCellByRcellAndCcell(sheet, cCell, tCell);
    }

    /**
     *  通过值来搜索对应的单元格
     * @param sheet 分页
     * @param value 值
     * @return      对应的单元格
     */
    private Cell getCellByValue(Sheet sheet, Object value) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                CellType cellType = cell.getCellTypeEnum();
                switch (cellType) {
                    case STRING:
                        String str = cell.getStringCellValue();
                        if (str.equals(value)) {
                            return cell;
                        }
                        break;
//                        System.out.println(cell.getRichStringCellValue().getString());
                    case NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
//                            System.out.println(cell.getDateCellValue());
                        } else {
                            double num = cell.getNumericCellValue();
                            if (isNumeric((value + ""))) {
                                if (value != "") {
                                    if (num == Double.valueOf(value + "")) {
                                        return cell;
                                    }
                                }
                            }
//                            System.out.println(cell.getNumericCellValue());
                        }
                        break;
                    case BOOLEAN:
//                        System.out.println(cell.getBooleanCellValue());
                        break;
                    case FORMULA:
//                        System.out.println(cell.getCellFormula());
                        String formula = cell.getCellFormula();
                        if (formula.equals(value)) {
                            return cell;
                        }
                        break;
                    case BLANK:
//                        System.out.println();
                        break;
                    default:
//                        System.out.println();
                }
            }
        }
        return null;
    }

    /**
     * 判读字符串是否是数字
     * @param str   字符串
     * @return  返回boolean值
     */
    private boolean isNumeric(String str) {
        Pattern pattern = Pattern.compile("[0-9]*");
        Matcher isNum = pattern.matcher(str);
        return isNum.matches();
    }

    /**
     * 搜索单元格，通过两个单元个横纵坐标来获取对应的单元格
     * @param sheet         分页
     * @param customer      客户
     * @param title         标题
     * @return              对应的单元格
     */
    private Cell selectCellByRcellAndCcell(Sheet sheet, Cell customer, Cell title) {
        int row = customer.getRowIndex();
        int col = title.getColumnIndex();
        Row row1 = sheet.getRow(row);
        return row1.getCell(col);
    }


    /**
     * 搜索单元格通过三个值，获得对应的单元格
     * @param sheet         分页
     * @param customer      客户
     * @param title         标题
     * @param year          年份
     * @return              对应的单元格
     */
    public Cell selectCellByCustomerTitleYear(XSSFSheet sheet, String customer, String title, String year) {
        List<Cell> years = selectCellsByValue(sheet, year + "年");
        Row row = sheet.getRow(5);
        int iCol = 0;
        for (Cell cell : years) {
            iCol = cell.getColumnIndex();
            String str = row.getCell(iCol).getStringCellValue();
            if (str.equals(title)) {
                break;
            }
        }
        Cell cell = getCellByValue(sheet, customer);
        if (cell != null) {
            int iRow = cell.getRowIndex();
            return sheet.getRow(iRow).createCell(iCol);
        }else{
            return null;
        }

    }

    /**
     * 搜索对应值的list单元格集合
     * @param sheet     分页
     * @param value     值
     * @return          对应的单元格集合
     */
    private List<Cell> selectCellsByValue(XSSFSheet sheet, Object value) {
        List<Cell> cells = new ArrayList<>();
        for (Row row : sheet) {
            for (Cell cell : row) {
                CellType cellType = cell.getCellTypeEnum();
                switch (cellType) {
                    case STRING:
                        String str = cell.getStringCellValue();
                        if (str.equals(value)) {
                            cells.add(cell);
                        }
                        break;
//                        System.out.println(cell.getRichStringCellValue().getString());
                    case NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
//                            System.out.println(cell.getDateCellValue());
                        } else {
                            double num = cell.getNumericCellValue();
                            if (isNumeric((value + ""))) {
                                if (num == Double.valueOf(value + "")) {
                                    cells.add(cell);
                                }
                            }
//                            System.out.println(cell.getNumericCellValue());
                        }
                        break;
                    case BOOLEAN:
//                        System.out.println(cell.getBooleanCellValue());
                        break;
                    case FORMULA:
//                        System.out.println(cell.getCellFormula());
                        String formula = cell.getCellFormula();
                        if (formula.equals(value)) {
                            cells.add(cell);
                        }
                        break;
                    case BLANK:
//                        System.out.println();
                        break;
                    default:
//                        System.out.println();
                }
//                cell.setCellType(CellType.STRING);
//                if (cell.getStringCellValue().equals(value)){
//                    cells.add(cell);
//                }
            }
        }
        return cells;
    }
}
