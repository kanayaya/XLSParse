import com.kanayaya.XLSParse.InnerClassImplementation.XLSTableParser;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;

class XLSTableParserTest {

    @Test
    void fromSheet() throws IOException {
        InputStream xlsStream = new BufferedInputStream(getClass().getResourceAsStream("/test.xlsx"));
        XSSFWorkbook book = new XSSFWorkbook(xlsStream);

        XLSTableParser.fromSheet(book.getSheetName(0))
                .findRowThat(row -> row.getCell(row.getFirstCellNum()).getStringCellValue().contains("title 1"))
                .thenSkip(1)
                .endIf(row -> ! isNumeric(new DataFormatter().formatCellValue(row.getCell(row.getFirstCellNum()))))
                .getEntityFrom(() -> new LinkedHashMap<String, String>())

                .thenForColumnValue((dto, cell) -> dto.put(cell, cell))
                .thenForColumnValue((dto, cell) -> dto.put(cell, cell))
                .thenPutTo(new ArrayList<>())
                .parse(book);
    }
    private boolean isNumeric(String s) {
        try {
            Double.parseDouble(s);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }
}