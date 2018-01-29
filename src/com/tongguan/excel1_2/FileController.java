package com.tongguan.excel1_2;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class FileController {
    /**
     * 构造方法
     */
    FileController() {
    }

    /**
     * 获取文件返回XSSFWorkbook对象
     * @param filePath 文件路径
     * @return 返回操作对象
     * @throws InvalidFormatException
     */
    public XSSFWorkbook getFile(String filePath) {
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(new FileInputStream(filePath));
        } catch (IOException e) {
//            e.printStackTrace();
            System.out.println("文件读取失败："+e);
        } catch (InvalidFormatException e) {
//            e.printStackTrace();
            System.out.println("文件读取失败："+e);
        }
        return (XSSFWorkbook) workbook;
    }

    /**
     * 创建文件
     * @param filePath  文件位置
     * @param workbook  哪个工作薄
     */
    public void createFile(String filePath,XSSFWorkbook workbook) {
        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream(filePath);
            workbook.write(fileOut);
        } catch (IOException e) {
//            e.printStackTrace();
            System.out.println("文件写入失败:"+e.toString());
        }finally {
            try {
                assert fileOut != null;
                fileOut.close();
            } catch (IOException e) {
//                e.printStackTrace();
                System.out.println("文件关闭失败:"+e.toString());
            }
        }
    }

    /**
     * 复制文件
     * @param fromPath  复制的源文件
     * @param toPath    复制到哪
     * @return  返回复制文件的工作簿
     */
    public XSSFWorkbook copyFile(String fromPath,String toPath){
        File fromFile = new File(fromPath);
        File toFile = new File(toPath);
        try {
            copyFile(fromFile,toFile);
        } catch (IOException e) {
//            e.printStackTrace();
            System.out.println("文件复制失败！"+e);
        }
        return getFile(toPath);
    }

    /**
     * 复制文件
     * @param fromFile 复制的源文件file对象
     * @param toFile   file对象
     * @throws IOException
     */
    private void copyFile(File fromFile, File toFile) throws IOException{
        FileInputStream ins = new FileInputStream(fromFile);
        FileOutputStream out = new FileOutputStream(toFile);
        byte[] b = new byte[1024];
        int n=0;
        while((n=ins.read(b))!=-1){
            out.write(b, 0, n);
        }
        ins.close();
        out.close();
    }
}
