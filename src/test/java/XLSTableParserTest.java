import com.kanayaya.XLSParse.InnerClassImplementation.XLSTableParser;
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

                .thenForNextColumnStringified((dto, s) -> dto.put(s, s))
                .thenForNextColumn((dto, cell) -> dto.put(cell.getRawValue(), Integer.toString(Double.valueOf(cell.getNumericCellValue()).intValue())))
                .thenPutInto(System.out::println)

                .thenContinueSameSheet()
                .findRowWhereCell(0).isNotNull().and().isNotEmpty().and().stringValueContains("title 1")
                .thenSkip(1)
                .endIfCell(2).isNull().or().isEmpty()
                .getEntityFrom(() -> new LinkedHashMap<String, String>())

                .thenForNextColumn((dto, cell) -> dto.put(cell.getStringCellValue(), cell.getStringCellValue()))
                .thenForNextColumn((dto, cell) -> dto.put(cell.getRawValue(), cell.getStringCellValue()))
                .thenPutInto(System.out::println)

                .parse(book);

        XLSTableParser.fromSheet(0)
                .findRowThat(row -> row.getCell(row.getFirstCellNum()).getStringCellValue().contains("title 1"))
                .thenSkip(1)
                .endIfCell(0).isNull().or().isEmpty().or().isNotNumeric()
                .getEntityFrom(() -> new LinkedHashMap<String, String>())

                .thenForNextColumnStringified((dto, s) -> dto.put(s, s))
                .thenForNextColumn((dto, cell) -> dto.put(cell.getRawValue(), Integer.toString(Double.valueOf(cell.getNumericCellValue()).intValue())))
                .thenPutInto(System.out::println)

                .thenRestartSameSheet()
                .findRowWhereCell(0).isNotNull().and().isNotEmpty().and().stringValueContains("title 1")
                .thenSkip(1)
                .endIfCell(2).isNull().or().isEmpty()
                .getEntityFrom(() -> new LinkedHashMap<String, String>())

                .thenForColumn(0, (dto, cell) -> dto.put(Double.toString(cell.getNumericCellValue()), Double.toString(cell.getNumericCellValue())))
                .thenForColumn(1, (dto, cell) -> dto.put(cell.getRawValue(), Double.toString(cell.getNumericCellValue())))
                .thenPutInto(System.out::println)

                .thenFromSheet(0)
                .findRowWhereCell(0).isNotNull().and().isNotEmpty().and().stringValueContains("title 1")
                .thenSkip(1)
                .endIfCell(2).isNull().or().isEmpty()
                .getEntityFrom(() -> new LinkedHashMap<String, String>())

                .thenForColumn(0, (dto, cell) -> dto.put(Double.toString(cell.getNumericCellValue()), Double.toString(cell.getNumericCellValue())))
                .thenForColumn(1, (dto, cell) -> dto.put(cell.getRawValue(), Double.toString(cell.getNumericCellValue())))
                .thenPutInto(System.out::println)

                .thenFromSheet(book.getSheetName(0))
                .findRowWhereCell(0).isNotNull().and().isNotEmpty().and().stringValueContains("title 1")
                .noSkip()
                .endIfCell(0).isNull().or().isEmpty().or().isNotString().orOtherCell(1).isNotNull()
                .getEntityFrom(() -> new LinkedHashMap<String, String>())

                .thenForColumn(0, (dto, cell) -> dto.put(cell.getStringCellValue(), cell.getStringCellValue()))
                .thenPutInto(System.out::println)

                .parse(book);
    }
}