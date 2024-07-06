package com.kanayaya.XLSParse.InnerClassImplementation;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;
import java.util.Objects;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Predicate;
import java.util.stream.IntStream;

@Slf4j
class TableFiller<T> {
    private final Function<XSSFWorkbook, XSSFSheet> sheetGetter;
    private final List<UncheckedBiConsumer<T, String>> columnFillers;
    private final UncheckedSupplier<T> getter;
    private final Predicate<XSSFRow> rowFilter;
    private final Predicate<XSSFRow> stopIf;
    private final int skip;
    private final Consumer<T> filler;

    TableFiller(
            Function<XSSFWorkbook, XSSFSheet> sheetGetter,
            List<UncheckedBiConsumer<T, String>> columnFillers,
            UncheckedSupplier<T> getter,
            Predicate<XSSFRow> rowFilter,
            Predicate<XSSFRow> stopIf,
            int skip,
            Consumer<T> filler) {
        this.sheetGetter = sheetGetter;
        this.columnFillers = columnFillers;
        this.getter = getter;
        this.rowFilter = rowFilter;
        this.stopIf = stopIf;
        this.skip = skip;
        this.filler = filler;
    }

    void fillFrom(XSSFWorkbook book) {
        XSSFSheet sheet = sheetGetter.apply(book);
        IntStream.range(0, sheet.getLastRowNum() + 1)
                .mapToObj(sheet::getRow)
                .filter(Objects::nonNull)
                .dropWhile(rowFilter.negate())
                .skip(skip)
                .takeWhile(stopIf.negate())
                .forEach(row ->  {
                    T data = getter.getUnchecked();
                    if (columnFillers.size() > row.getLastCellNum()) {
                        throw new IllegalStateException(String.format("Недостаточно ячеек в ряду. Количество обработчиков для каждой -- %d, а номер последней ячейки -- %d", columnFillers.size(), row.getLastCellNum()));
                    } else if (columnFillers.size() < row.getLastCellNum()) {
                        log.warn(String.format("Количество ячеек в ряду (%d) больше количества обработчиков для каждой (%d), вы уверены что указали все обработчики?", row.getLastCellNum(), columnFillers.size()));
                    }
                    for (int i = row.getFirstCellNum(); i < columnFillers.size(); i++) {
                        XSSFCell cell = row.getCell(i);
                        if (cell == null) {throw new NullPointerException(String.format("Ячейка %d ряда %d листа %s не существует", i, row.getRowNum(), sheet.getSheetName()));}
                        columnFillers.get(i - row.getFirstCellNum()).acceptUnchecked(data, new DataFormatter().formatCellValue(cell));
                    }
                    filler.accept(data);
                });
    }
}