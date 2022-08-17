package com.localhost.report;

import java.io.*;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.Instant;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.function.Consumer;
import java.util.function.Function;

import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.validation.BindingResult;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;

import static java.time.temporal.ChronoUnit.*;

@Controller
public class RecordController {

    @GetMapping
    public String read(@ModelAttribute Record record,
                       BindingResult errors,
                       Model model,
                       @Value("${file.path}") String filePath)
        throws IOException {

        var file = new FileInputStream(filePath);
        var workbook = new XSSFWorkbook(file);
        var sheet = workbook.getSheetAt(0);

        var dd = new ArrayList<Record>();
        for (Row row : sheet) {
            if (row.getRowNum() == 0) {
                continue;
            }
            dd.add(Record.builder()
                .name(row.getCell(0).toString())
                .ip(row.getCell(1).toString())
                .user(row.getCell(2).toString())
                .password(row.getCell(3).toString())
                .ipmiIp(row.getCell(4).toString())
                .ipmiUser(row.getCell(5).toString())
                .ipmiPassword(row.getCell(6).toString())
                .build()
            );
        }
        model.addAttribute("data", dd);
        return "index.htm";
    }

    public static void main(String[] args) throws IOException {
        create();
    }

    @GetMapping("create")
    public static void create() {
        File excel = new File("/Users/brunotaboada/IdeaProjects/report/src/main/resources/excel.xlsx");
        try (
            InputStream fis = new FileInputStream(excel);
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
        ) {
            XSSFSheet sheet = workbook.getSheetAt(0);
            int totalNumberOfRows = sheet.getLastRowNum();
            SimpleDateFormat formatter = new SimpleDateFormat("M/d/yyyy");
            DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("M/d/yyyy")
                    .withZone(ZoneId.systemDefault());
            Calendar c = Calendar.getInstance();
            int rowNum = 0;
            boolean skipFirstLine = true;
            for(int i = 0; i <= totalNumberOfRows; i++){
                if(skipFirstLine){
                    skipFirstLine = false;
                    rowNum++;
                    continue;
                }
                XSSFRow row = sheet.getRow(rowNum);
                if(row == null){
                    continue;
                }
                String strDate = row.getCell(0).getStringCellValue();
                Date date = formatter.parse(strDate);
                c.setTime(date);
                int datOfWeek = c.get(Calendar.DAY_OF_WEEK);
                if(datOfWeek != Calendar.FRIDAY){
                    rowNum++;
                    continue;
                }
                sheet.shiftRows(row.getRowNum()+1, sheet.getLastRowNum() + 2, 2, true, true);
                XSSFRow newCell1 = sheet.createRow(row.getRowNum()+1);
                XSSFRow newCell2 = sheet.createRow(row.getRowNum()+2);
                rowNum += 3;
                newCell1.copyRowFrom(row, new CellCopyPolicy());
                newCell2.copyRowFrom(row, new CellCopyPolicy());
                Instant newDate = date.toInstant().plus(1, DAYS);
                Instant newDate2 = date.toInstant().plus(2, DAYS);
                newCell1.getCell(0).setCellValue(dateTimeFormatter.format(newDate));
                newCell2.getCell(0).setCellValue(dateTimeFormatter.format(newDate2));
                newCell1.getCell(1).setCellValue(dateTimeFormatter.format(newDate));
                newCell2.getCell(1).setCellValue(dateTimeFormatter.format(newDate2));
            }
            //Modify same file
            //OutputStream os = new FileOutputStream(excel);
            //Create new file
            OutputStream os = new FileOutputStream("./excel.xlsx");
            workbook.write(os);
        } catch (IOException | ParseException e) {
            throw new RuntimeException(e);
        }
    }

}