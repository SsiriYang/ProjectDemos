package com.ys.demo.freemark.controller;


import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.apache.tomcat.util.http.fileupload.FileUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;
import springfox.documentation.swagger2.annotations.EnableSwagger2;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Api("报告导出")
@EnableSwagger2
@RestController
public class ReportExportController {
    private static final Logger logger = LoggerFactory.getLogger(ReportExportController.class);

    /**
     * 导出报告 Word
     * @throws IOException
     * @return
     */
    @GetMapping("/export")
    @ApiOperation(value="导出报告", httpMethod = "GET",produces="application/json",notes = "导出报告doc")
    public void exportDoc(HttpServletResponse response) throws IOException {

        Configuration configuration = new Configuration();
        configuration.setDefaultEncoding("utf-8");
        // 模板存放路径
        configuration.setClassForTemplateLoading(this.getClass(), "/reportTemplates");
        // 获取模板文件
        Template template = configuration.getTemplate("test.ftl");
        // 获取数据

        List<HashMap<String,String>> assetTypeList = new ArrayList<>();
        HashMap<String,String> map  = new HashMap();
        map.put("类型a","1");
        map.put("类型b","2");
        assetTypeList.add(map);
        Map<String, Object> resultMap = new HashMap<>();
        resultMap.put("assetTypeList", assetTypeList);
        File outFile = new File("Demo.docx");
        Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile), "UTF-8"));
        try {
            template.process(resultMap, out);
            out.flush();
            out.close();
        } catch (TemplateException e) {
            e.printStackTrace();
        }
        if (outFile.length() == 0L) {
            logger.error("报告导出失败");
        }
        // 以流的形式下载文件。
        InputStream fis = new BufferedInputStream(new FileInputStream(outFile));
        byte[] buffer = new byte[fis.available()];
        fis.read(buffer);
        fis.close();
        // 清空response
        response.reset();
        // 设置response的Header
        String fileName = "test";
        response.addHeader("Content-Disposition", "attachment;filename=" + new String(fileName.getBytes()));
        response.addHeader("Content-Length", "" + outFile.length());
        OutputStream toClient = new BufferedOutputStream(response.getOutputStream());
        response.setContentType("application/octet-stream");
        toClient.write(buffer);
        toClient.flush();
        toClient.close();
    }

}
