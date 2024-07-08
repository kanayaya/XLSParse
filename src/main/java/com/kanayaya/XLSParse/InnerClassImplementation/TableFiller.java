package com.kanayaya.XLSParse.InnerClassImplementation;

import lombok.Getter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;
import java.util.Objects;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.Function;
import java.util.function.Predicate;
import java.util.stream.IntStream;

@Slf4j
class TableFiller<T> {
    private final AtomicBoolean continueNext = new AtomicBoolean(false);
    private final AtomicInteger rowCounter = new AtomicInteger();
    @Getter
    private final Function<XSSFWorkbook, XSSFSheet> sheetGetter;
    private final List<UncheckedBiConsumer<T, XSSFCell>> columnFillers;
    private final UncheckedSupplier<T> getter;
    private final Predicate<XSSFRow> startIf;
    private final Predicate<XSSFRow> stopIf;
    private final int skip;
    private final UncheckedConsumer<T> filler;

    TableFiller(
            Function<XSSFWorkbook, XSSFSheet> sheetGetter,
            List<UncheckedBiConsumer<T, XSSFCell>> columnFillers,
            UncheckedSupplier<T> getter,
            Predicate<XSSFRow> rowFilter,
            Predicate<XSSFRow> stopIf,
            int skip,
            UncheckedConsumer<T> filler) {
        this.sheetGetter = sheetGetter;
        this.columnFillers = columnFillers;
        this.getter = getter;
        this.startIf = rowFilter;
        this.stopIf = stopIf;
        this.skip = skip;
        this.filler = filler;
    }
    void continueWhereEnded() {
        continueNext.set(true);
    }

    int fillFrom(XSSFWorkbook book, int start) {
        AtomicBoolean logToWarn = new AtomicBoolean(true);
        log.info("Начинаем парсинг XLS со строки " + start);
        XSSFSheet sheet = sheetGetter.apply(book);
        if (start > sheet.getLastRowNum()) throw new IllegalArgumentException(String.format("Стартовый ряд (%d) не может быть больше максимального количества рядов на листе (%d)", start, sheet.getLastRowNum()));
        IntStream.range(start, sheet.getLastRowNum() + 1)
                .peek(rowCounter::set)
                .mapToObj(sheet::getRow)
                .filter(Objects::nonNull)
                .dropWhile(startIf.negate())
                .skip(skip)
                .takeWhile(stopIf.negate())
                .forEach(row ->  {
                    T data = getter.get();
                    if (columnFillers.size() > row.getLastCellNum()) {
                        throw new IllegalStateException(String.format("Недостаточно ячеек в ряду. Количество обработчиков для каждой -- %d, а номер последней ячейки -- %d", columnFillers.size(), row.getLastCellNum()));
                    } else if (columnFillers.size() < row.getLastCellNum()) {
                        String message = String.format("Количество ячеек в ряду (%d) больше количества обработчиков для каждой (%d), вы уверены что указали все обработчики?", row.getLastCellNum(), columnFillers.size());
                        if (logToWarn.get()) {
                            log.warn(message);
                            logToWarn.set(false);
                        } else {
                            log.debug(message);
                        }
                    }
                    for (int i = row.getFirstCellNum(); i < columnFillers.size(); i++) {
                        XSSFCell cell = row.getCell(i);
                        if (cell == null) {throw new NullPointerException(String.format("Ячейка %d ряда %d листа \"%s\" не существует", i, row.getRowNum(), sheet.getSheetName()));}
                        try {
                            columnFillers.get(i - row.getFirstCellNum()).accept(data, cell);
                        } catch (RuntimeException e) {
                            if (e.getCause() instanceof IllegalStateException) throw new IllegalStateException("Несовпадение типов поля и запрашиваемого значения. См. причину: " + e.getCause().getMessage(), e.getCause());
                            else throw e;
                        }
                    }
                    filler.accept(data);
                });
        return continueNext.get()? rowCounter.get() : 0;
    }
}