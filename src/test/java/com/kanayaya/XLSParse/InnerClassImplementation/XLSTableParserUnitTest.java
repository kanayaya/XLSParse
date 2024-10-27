package com.kanayaya.XLSParse.InnerClassImplementation;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.util.LinkedHashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;
@ExtendWith(MockitoExtension.class)
class XLSTableParserUnitTest {
    XLSTableParser.StartConditionGetter startConditionGetter = XLSTableParser.fromSheet("mock sheet name");




    @Test
    void testStartConditionGetter() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.getSheet("Лист 1");
        int headerRowIndex = -1;
        for (int i = sheet.getFirstRowNum(); i < sheet.getLastRowNum() + 1; i++) {
            XSSFRow row = sheet.getRow(i);
            if (row.getCell(row.getFirstCellNum()).getCellType().equals(CellType.STRING)
                    && row.getCell(row.getFirstCellNum()).getStringCellValue().contains("Заголовок таблицы")) {
                headerRowIndex = i;
                break;
            }
        }
        if (headerRowIndex >= 0) {
            int i = headerRowIndex + 2;
            while (i <= sheet.getLastRowNum()) {
                XSSFRow row = sheet.getRow(i);
                if (row == null || ! row.getCell(0).getCellType().equals(CellType.NUMERIC)) break;
                Map<String, String> dto = new LinkedHashMap<>();
                String val = new DataFormatter().formatCellValue(row.getCell(0));
                dto.put(val, val);
                String val2 = Integer.toString(Double.valueOf(row.getCell(1).getNumericCellValue()).intValue());
                dto.put(row.getCell(1).getRawValue(), val2);
                System.out.println(dto);
            }
        }

        assertTrue(startConditionGetter.findRowThat(row->true) instanceof XLSTableParser.Skipper);
    }

}