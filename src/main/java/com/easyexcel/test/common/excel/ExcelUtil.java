package com.easyexcel.test.common.excel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.easyexcel.test.common.excel.exception.ExcelException;
import com.easyexcel.test.common.excel.listener.ExcelListener;
import com.easyexcel.test.common.excel.utils.DateUtils;
import com.easyexcel.test.common.excel.utils.MyBeanCopy;
import com.google.common.collect.Lists;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * ExcelUtil
 * 基于easyExcel的开源框架，poi版本3.17
 * BeanCopy ExcelException 属于自定义数据，属于可自定义依赖
 * 工具类尽可能还是需要减少对其他java的包的依赖
 * @author parker
 */
@Slf4j
public final class ExcelUtil {
    /**
     * 私有化构造方法
     */
    private ExcelUtil(){}

    /**
     * 读取 Excel(多个 sheet)
     * 将多sheet合并成一个list数据集，通过自定义ExcelReader继承AnalysisEventListener
     * 重写invoke doAfterAllAnalysed方法
     * getExtendsBeanList 主要是做Bean的属性拷贝 ，可以通过ExcelReader中添加的数据集直接获取
     * @param excel    文件
     * @param rowModel 实体类映射，继承 BaseRowModel 类
     * @return Excel 数据 list
     */
    public static <T> List<T> readExcel(MultipartFile excel,Class<T>  rowModel) throws ExcelException {
        return readExcel(excel, rowModel, null, 1);
    }

    /**
     * 读取某个 sheet 的 Excel
     * @param excel    文件
     * @param rowModel 实体类映射，继承 BaseRowModel 类
     * @param sheetName  sheet 的序号 从1开始
     * @return Excel 数据 list
     */
    public static <T> List<T> readExcel(MultipartFile excel, Class<T>  rowModel, String sheetName)  throws ExcelException{
        return readExcel(excel, rowModel, sheetName, 1);
    }

    /**
     * 读取某个 sheet 的 Excel
     * @param excel       文件
     * @param rowModel    实体类映射，继承 BaseRowModel 类
     * @param sheetName     sheet 的序号 从1开始
     * @param headLineNum 表头行数，默认为1
     * @return Excel 数据 list
     */
    public static <T> List<T> readExcel(MultipartFile excel, Class<T>  rowModel, String sheetName,
                                                             int headLineNum) throws ExcelException {
        ExcelListener excelListener = new ExcelListener();

        InputStream inputStream = null;
        try{
            if(null != excel){
                inputStream = excel.getInputStream();
            }
        }catch (IOException e){
            log.error(e.getMessage(),e);
        }
        if(null == inputStream){
            return Lists.newArrayList();
        }

        ExcelReader excelReader = EasyExcel.read(inputStream, rowModel, excelListener).build();
        if (excelReader == null) {
            return Lists.newArrayList();
        }
        ReadSheet readSheet;
        if(StringUtils.isEmpty(sheetName)){
            readSheet = EasyExcel.readSheet().build();
        }else{
            readSheet = EasyExcel.readSheet(sheetName).build();
        }
        readSheet.setHeadRowNumber(headLineNum);
        excelReader.read(readSheet);
        // 这里千万别忘记关闭，读的时候会创建临时文件，到时磁盘会崩的
        excelReader.finish();

        return getExtendsBeanList(excelListener.getDataList(),rowModel);
    }

    /**
     * 导出 Excel ：一个 sheet，带表头
     * 自定义WriterHandler 可以定制行列数据进行灵活化操作
     * @param response  HttpServletResponse
     * @param list      数据 list，每个元素为一个 BaseRowModel
     * @param fileName  导出的文件名
     * @param sheetName 导入文件的 sheet 名
     */
    public static <T> void writeExcel(HttpServletResponse response, List<T> list,
                                                           String fileName, String sheetName,
                                                           Class<T> classType, ExcelTypeEnum excelTypeEnum) throws ExcelException{
        if(sheetName == null || "".equals(sheetName)){
            sheetName = "sheet1";
        }

        fileName = fileName+ DateUtils.getDate("yyyyMMddHHmmss");
        OutputStream outputStream = getOutputStream(fileName, response, excelTypeEnum);
        ExcelWriter excelWriter = EasyExcel.write(outputStream, classType).build();
        WriteSheet writeSheet = EasyExcel.writerSheet(1, sheetName).build();
        writeSheet.setRelativeHeadRowIndex(0);
        excelWriter.write(list,writeSheet);

        // 关闭流
        try {
            excelWriter.finish();
            outputStream.flush();
            outputStream.close();
        }catch (Exception e){
            log.error(e.getMessage(),e);
        }
    }


    /**
     * 导出文件时为Writer生成OutputStream
     */
    private static OutputStream getOutputStream(String fileName, HttpServletResponse response, ExcelTypeEnum excelTypeEnum) throws ExcelException{
        //创建本地文件
        String filePath = fileName + excelTypeEnum.getValue();
        try {
            fileName = new String(filePath.getBytes(), "ISO-8859-1");
            response.addHeader("Content-Disposition", "filename=" + fileName);
            return response.getOutputStream();
        } catch (IOException e) {
            throw new ExcelException("创建文件失败！");
        }
    }

    /**
     * 返回 ExcelReader
     * @param excel         需要解析的 Excel 文件
     * @param excelListener new ExcelListener()
     */
    private static ExcelReader getReader(MultipartFile excel,
                                         ExcelListener excelListener) throws ExcelException{
        String fileName = excel.getOriginalFilename();
        if (fileName == null ) {
            throw new ExcelException("文件格式错误！");
        }
        if (!fileName.toLowerCase().endsWith(ExcelTypeEnum.XLS.getValue()) && !fileName.toLowerCase().endsWith(ExcelTypeEnum.XLSX.getValue())) {
            throw new ExcelException("文件格式错误！");
        }
        InputStream inputStream;
        try {
            inputStream = excel.getInputStream();
            return EasyExcel.read(inputStream, excelListener).build();
        } catch (IOException e) {
            //do something
            log.error(e.getMessage(),e);
        }
        return null;
    }

    /**
     * 利用BeanCopy转换list
     */
    public static <T> List<T> getExtendsBeanList(List<?> list,Class<T> typeClazz){
        return MyBeanCopy.convert(list,typeClazz);
    }
}