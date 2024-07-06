package com.kanayaya.XLSParse.InnerClassImplementation;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.function.Function;
import java.util.function.Predicate;

public class XLSTableParser {
    private final List<TableFiller<?>> fillers = new ArrayList<>();
    public static StartCondition fromSheet(String sheetName) {
        return new StartCondition((workbook) -> workbook.getSheet(sheetName));
    }

    private XLSTableParser(TableFiller<?> filler) {
        fillers.add(filler);
    }
    public StartCondition thenFromSheet(String sheetName) {
        return new StartCondition((workbook) -> workbook.getSheet(sheetName));
    }
    public void parse(XSSFWorkbook book) {
        for (TableFiller<?> filler : fillers) {
            filler.fillFrom(book);
        }
    }

    public static final class StartCondition {

        private final Function<XSSFWorkbook, XSSFSheet> sheetGetter;
        private StartCondition(Function<XSSFWorkbook, XSSFSheet> sheetGetter) {
            this.sheetGetter = sheetGetter;
        }
        public Skipper findRowThat(Predicate<XSSFRow> filter) {
            return new Skipper(sheetGetter, filter);
        }


    }
    public static final class Skipper {

        private final Function<XSSFWorkbook, XSSFSheet> sheetGetter;
        private final Predicate<XSSFRow> filter;

        private Skipper(Function<XSSFWorkbook, XSSFSheet> sheetGetter, Predicate<XSSFRow> filter) {
            this.sheetGetter = sheetGetter;
            this.filter = filter;
        }
        public EndCondition thenSkip(int skip) {
            return new EndCondition(sheetGetter, filter, skip);
        }
    }
    public static final class EndCondition {
        private final Function<XSSFWorkbook, XSSFSheet> sheetGetter;
        private final Predicate<XSSFRow> filter;
        private final int skip;

        private EndCondition(Function<XSSFWorkbook, XSSFSheet> sheetGetter, Predicate<XSSFRow> filter, int skip) {
            this.sheetGetter = sheetGetter;
            this.filter = filter;
            this.skip = skip;
        }
        public EntityGetter endIf(Predicate<XSSFRow> rowDecliner) {
            return new EntityGetter(sheetGetter, filter, skip, rowDecliner);
        }
    }
    public static final class EntityGetter {

        private final Function<XSSFWorkbook, XSSFSheet> sheetGetter;
        private final Predicate<XSSFRow> filter;
        private final int skip;
        private final Predicate<XSSFRow> rowDecliner;

        private EntityGetter(Function<XSSFWorkbook, XSSFSheet> sheetGetter, Predicate<XSSFRow> filter, int skip, Predicate<XSSFRow> rowDecliner) {
            this.sheetGetter = sheetGetter;
            this.filter = filter;
            this.skip = skip;
            this.rowDecliner = rowDecliner;
        }
        public <T> EntityFiller<T> getEntityFrom(UncheckedSupplier<T> generator) {
            return new EntityFiller<>(sheetGetter, filter, skip, rowDecliner, generator);
        }
    }

    public static final class EntityFiller<T> {

        private final Function<XSSFWorkbook, XSSFSheet> sheetGetter;
        private final Predicate<XSSFRow> filter;
        private final int skip;
        private final Predicate<XSSFRow> rowDecliner;
        private final UncheckedSupplier<T> generator;
        private final List<UncheckedBiConsumer<T, String>> columnFillers = new ArrayList<>();

        private EntityFiller(Function<XSSFWorkbook, XSSFSheet> sheetGetter, Predicate<XSSFRow> filter, int skip, Predicate<XSSFRow> rowDecliner, UncheckedSupplier<T> generator) {
            this.sheetGetter = sheetGetter;
            this.filter = filter;
            this.skip = skip;
            this.rowDecliner = rowDecliner;
            this.generator = generator;
        }
        public EntityFiller<T> thenForColumnValue(UncheckedBiConsumer<T, String> filler) {
            columnFillers.add(filler);
            return this;
        }
        public XLSTableParser thenPutTo(Collection<? super T> collection) {
            return new XLSTableParser(new TableFiller<>(sheetGetter, columnFillers, generator, filter, rowDecliner, skip, collection::add));
        }
    }
}