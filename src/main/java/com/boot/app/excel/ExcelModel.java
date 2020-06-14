package com.boot.app.excel;

import java.util.Collections;
import java.util.List;
import java.util.function.Consumer;

import org.apache.poi.ss.usermodel.Workbook;

import lombok.Getter;

@Getter
public class ExcelModel<T> {
    
    private String fileName;
    private String[] headerKeys;
    private String[] headerNames;
    private final List<T> rows;
    private Consumer<Workbook> customizer;
    
    private ExcelModel(List<T> rows) {
        this.rows = rows;
    }
    
    public static <T> ExcelModel<T> withEmptyRows() {
        return of(Collections.emptyList());
    }
    
    public static <T> ExcelModel<T> of(List<T> rows) {
        return new ExcelModel<>(rows);
    }
    
    public ExcelModel<T> setHeaderKeys(String ...headerKeys) {
        this.headerKeys = headerKeys;
        return this;
    }
    
    public ExcelModel<T> setHeaderNames(String ...headerNames) {
        this.headerNames = headerNames;
        return this;
    }
    
    public ExcelModel<T> setFileName(String fileName) {
        this.fileName = fileName;
        return this;
    }
    
    public ExcelModel<T> setCustomizer(Consumer<Workbook> customizer) {
        this.customizer = customizer;
        return this;
    }
}
