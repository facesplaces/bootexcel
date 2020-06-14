package com.boot.app.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;

@Controller
public class ExcelController {
    
    @GetMapping("/")
    public String main() {
        return "ok";
    }

    @GetMapping("/form")
    public ExcelView<SampleVO> form(Model model) {
        model.addAttribute(
                ExcelView.EXCEL_MODEL, 
                ExcelModel.withEmptyRows()
                .setHeaderKeys("name", "value")
                .setHeaderNames("이름", "값")
                .setFileName("excel.xlsx")
                .setCustomizer(firstColumnToTextFormat()));
        return new ExcelView<>();
    }
    private Consumer<Workbook> firstColumnToTextFormat() {
        return workbook -> {
            DataFormat format = workbook.createDataFormat();
            CellStyle textStyle = workbook.createCellStyle();
            textStyle.setDataFormat(format.getFormat("@"));
            Sheet sheet = workbook.getSheetAt(0);
            sheet.setDefaultColumnStyle(0, textStyle);
        };
    }
    
    @GetMapping("/download")
    public ExcelView<SampleVO> download(Model model) {
        model.addAttribute(
            ExcelView.EXCEL_MODEL, 
            ExcelModel.of(getSampleRows())
                .setHeaderKeys("name", "value")
                .setHeaderNames("이름", "값")
                .setFileName("excel.xlsx"));
        return new ExcelView<>();
    }
    private List<SampleVO> getSampleRows() {
        List<SampleVO> rows = new ArrayList<>();
        for(int i=0;i<100000;i++) {
            rows.add(new SampleVO("name"+i ,"value"+i));
        }
        return rows;
    }
    
}
