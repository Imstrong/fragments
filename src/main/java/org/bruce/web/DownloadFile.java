package org.bruce.web;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.support.ExcelTypeEnum;
import org.apache.commons.io.FileUtils;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;
import java.util.Arrays;
import java.util.List;
import java.util.Locale;

@RestController
public class DownloadFile {


    private static final String CLUE_TEMPLATE_NAME = "hello.xlsx";
    private static final String TARGET_TEMPLATE_NAME = "导出了.xlsx";

    @GetMapping("download")
    public void download(HttpServletResponse response) {
        try (ServletOutputStream os = response.getOutputStream();
             InputStream is = this.getClass().getResourceAsStream("/" + CLUE_TEMPLATE_NAME)) {
            response.reset();
            response.setContentType("application/octet-stream");
            response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(TARGET_TEMPLATE_NAME, "utf-8"));
            response.setHeader("Cache-Control", "must-revalidate,post-check=0,pre-check=0");
            byte[] buf = new byte[1024];
            int len = 0;
            while ((len = is.read(buf)) > 0) {
                os.write(buf, 0, len);
                os.flush();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @PostMapping("import")
    public void importEmployee(@RequestPart("file") MultipartFile file) {
        String fileName = file.getOriginalFilename().trim();
        if (!Arrays.asList("xlsx", "xls").contains(fileName.substring(fileName.lastIndexOf('.') + 1))) {
            throw new RuntimeException("文件格式错误");
        }
        File tempFile = null;
        try {
            tempFile = tempSave(file.getInputStream());
            List<Object> objects = EasyExcel.read(tempFile)
                    .excelType(ExcelTypeEnum.XLSX)
                    .sheet(0)
                    .headRowNumber(-1)
                    .head(Object.class)
                    .autoTrim(true)
                    .doReadSync();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (tempFile != null && tempFile.exists()) {
                    FileUtils.forceDelete(tempFile);
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
    private File tempSave(InputStream stream) {
        File result = null;
        try {
            String path = this.getClass().getResource("/").toURI().getPath();
            result = new File(path, String.format(Locale.ROOT, "/temp/employee_import_%s_%s.xlsx", "userCode", System.currentTimeMillis()));
            FileUtils.copyToFile(stream, result);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }
}
