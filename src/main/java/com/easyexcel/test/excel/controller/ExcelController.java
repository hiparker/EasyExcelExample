package com.easyexcel.test.excel.controller;

import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.util.CollectionUtils;
import com.easyexcel.test.common.excel.ExcelUtil;
import com.easyexcel.test.common.excel.exception.ExcelException;
import com.easyexcel.test.common.result.Result;
import com.easyexcel.test.excel.entity.Test;
import com.google.common.collect.Lists;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

import javax.servlet.http.HttpServletResponse;
import java.util.Iterator;
import java.util.List;

/**
 * Created Date by 2020/5/9 0009.
 *
 * @author Parker
 */
@Slf4j
@Controller
@RequestMapping("/excel")
public class ExcelController {

    @RequestMapping("/form")
    public String excelForm(){
        return "/excelForm";
    }


    /**
     * 导出
     * @param response
     */
    @ResponseBody
    @RequestMapping("/export")
    public Result excelExport(HttpServletResponse response){
        log.info("进入 导出Excel Controller");

        Result result = new Result();

        String fileName = "Excel测试";

        List<Test> list = Lists.newArrayList();
        Test excel1 = new Test();
        excel1.setName("周一");
        excel1.setCode("x001");
        excel1.setAge(28);
        excel1.setAmt(1563.21);
        excel1.setSex(1);
        Test excel2 = new Test();
        excel2.setName("崔二");
        excel2.setCode("x002");
        excel2.setAge(25);
        excel2.setAmt(554342.1);
        excel2.setSex(0);
        Test excel3 = new Test();
        excel3.setName("张三");
        excel3.setCode("x003");
        excel3.setAge(31);
        excel3.setAmt(4450.1);
        excel3.setSex(1);
        Test excel4 = new Test();
        excel4.setName("李四");
        excel4.setCode("x004");
        excel4.setAge(15);
        excel4.setAmt(0d);
        excel4.setSex(1);

        list.add(excel1);
        list.add(excel2);
        list.add(excel3);
        list.add(excel4);

        try {
            ExcelUtil.writeExcel(response,list,fileName,"没有设定sheet名称", Test.class,ExcelTypeEnum.XLSX);
            result.setMsg("导出成功！");
            return result;
        } catch (ExcelException e) {
            log.error(e.getMessage(),e);
            result.setMsg("导出失败 Message = " + e.getMessage());
        }
        return result;
    }


    @ResponseBody
    @RequestMapping("/import")
    public Result excelImport(MultipartHttpServletRequest request){
        log.info("进入 导入Excel Controller");
        Result result = new Result();

        Iterator<String> itr = request.getFileNames();
        String uploadedFile = itr.next();
        List<MultipartFile> files = request.getFiles(uploadedFile);
        if (CollectionUtils.isEmpty(files)) {
            return Result.error("请选择文件");
        }

        try {
            List<Test> tests = ExcelUtil.readExcel(files.get(0), Test.class);
            result.put("data",tests);
            result.setMsg("导入成功！");
            return result;
        } catch (ExcelException e) {
            log.error(e.getMessage(),e);
            result.setMsg("导入失败 Message = " + e.getMessage());
        }
        return result;
    }

    /**
     * 下载导入模板
     * @param response
     */
    @ResponseBody
    @RequestMapping("/import/template")
    public Result importTemplate(HttpServletResponse response){
        log.info("进入 下载导入模板 Controller");

        Result result = new Result();

        String fileName = "Excel测试-导入模板";
        List<Test> list = Lists.newArrayList();
        try {
            ExcelUtil.writeExcel(response,list,fileName,"没有设定sheet名称", Test.class,ExcelTypeEnum.XLSX);
            result.setMsg("导出成功！");
            return result;
        } catch (ExcelException e) {
            log.error(e.getMessage(),e);
            result.setMsg("导出失败 Message = " + e.getMessage());
        }
        return result;
    }

}
