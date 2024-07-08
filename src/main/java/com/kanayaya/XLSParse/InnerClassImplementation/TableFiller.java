package com.kanayaya.XLSParse.InnerClassImplementation;

import lombok.NonNull;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Objects;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.Function;
import java.util.function.Predicate;
import java.util.stream.IntStream;

/**
 * Структура, содержащая инструкции для парсинга таблицы и метод, совершающий парсинг
 * @param <T> Тип DTO, куда кладутся результаты парсинга
 */
@Slf4j
class TableFiller<T> {
    /**
     * Функция, возвращающая {@link XSSFSheet} для парсинга. Нужна для задания логики доставания листа из {@link XSSFWorkbook}
     */
    private final Function<XSSFWorkbook, XSSFSheet> sheetGetter;
    /**
     * Предикат, определяющий, с какого ряда таблицы начинать парсинг
     */
    private final Predicate<XSSFRow> startIf;
    /**
     * Количество рядов, пропускаемых перед началом парсинга таблицы.
     * Нужно если первые N рядов таблицы являются заголовками или их не нужно парсить
     */
    private final int skip;
    /**
     * Предикат, определяющий, на каком ряду таблицы закончить парсинг
     */
    private final Predicate<XSSFRow> stopIf;
    /**
     * Генератор новых DTO для наполнения данными парсинга. Генерируется новый DTO для каждого ряда таблицы.
     */
    private final UncheckedSupplier<T> getter;
    /**
     * Набор инструкций от первой ячейки, задающих метод парсинга каждого столбца каждого ряда
     * таблицы.
     */
    private final UncheckedBiConsumer<T, XSSFRow> columnFiller;
    /**
     * Нужен для того, чтобы складывать туда созданные и наполненные DTO
     */
    private final UncheckedConsumer<? super T> dtoConsumer;

    TableFiller(
            Function<XSSFWorkbook, XSSFSheet> sheetGetter,
            UncheckedBiConsumer<T, XSSFRow> columnFiller,
            UncheckedSupplier<T> getter,
            Predicate<XSSFRow> rowFilter,
            Predicate<XSSFRow> stopIf,
            int skip,
            UncheckedConsumer<? super T> filler) {
        this.sheetGetter = sheetGetter;
        this.columnFiller = columnFiller;
        this.getter = getter;
        this.startIf = rowFilter;
        this.stopIf = stopIf;
        this.skip = skip;
        this.dtoConsumer = filler;
    }

    /**
     * Метод для запуска парсинга таблицы.
     * <p>Собирает данные после сбора инструкций и парсит по ним выбранную книгу</p>
     * @param book Книга, в которой находится таблица
     * @param start Номер ряда, с которого начинается парсинг
     * @return Ноль для следующего парсера
     */
    int fillFrom(@NonNull XSSFWorkbook book, int start) {
        fillContinuing(book, start);
        return 0;
    }

    /**
     * Метод для запуска парсинга таблицы.
     * <p>Собирает данные после сбора инструкций и парсит по ним выбранную книгу</p>
     * @param book Книга, в которой находится таблица
     * @param start Номер ряда, с которого начинается парсинг
     * @return Номер строки, на которой закончился парсинг
     */
    int fillContinuing(@NonNull XSSFWorkbook book, int start) {
        final AtomicInteger rowCounter = new AtomicInteger(start);
        XSSFSheet sheet = sheetGetter.apply(book);
        log.info(String.format("Начинаем парсинг XLS-листа \"%s\" со строки %d", sheet.getSheetName(), start));
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
                    columnFiller.accept(data, row);
                    dtoConsumer.accept(data);
                });
        return rowCounter.get();
    }
}
  