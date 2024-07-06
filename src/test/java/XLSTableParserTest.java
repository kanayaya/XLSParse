
import com.kanayaya.XLSParse.InnerClassImplementation.XLSTableParser;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.LinkedHashMap;

class XLSTableParserTest {

    @Test
    void fromSheet() throws IOException {
        InputStream xlsStream = new BufferedInputStream(getClass().getResourceAsStream("/test.xlsx"));
        XSSFWorkbook book = new XSSFWorkbook(xlsStream);

        XLSTableParser.fromSheet(book.getSheetName(0))
                .findRowThat(row -> row.getCell(row.getFirstCellNum()).getStringCellValue().contains("title 1"))
                .thenSkip(1)
                .endIfCell(0).isNull().or().isEmpty().or().isNotNumeric()
                .getEntityFrom(() -> new LinkedHashMap<String, String>())

                .thenForColumnStringifiedValue((dto, s) -> dto.put(s, s))
                .thenForColumn((dto, cell) -> dto.put(cell.getRawValue(), Integer.toString(Double.valueOf(cell.getNumericCellValue()).intValue())))
                .thenPutInto(System.out::println)

                .thenContinueSameSheet()
                .findRowThat(row -> new DataFormatter().formatCellValue(row.getCell(row.getFirstCellNum())).contains("title 1"))
                .thenSkip(1)
                .endIfCell(2).isNull().or().isEmpty()
                .getEntityFrom(() -> new LinkedHashMap<String, String>())

                .thenForColumn((dto, cell) -> dto.put(cell.getStringCellValue(), cell.getStringCellValue()))
                .thenForColumn((dto, cell) -> dto.put(cell.getRawValue(), cell.getStringCellValue()))
                .thenPutInto(System.out::println)

                .parse(book);
    }
}