package com.kanayaya.XLSParse.InnerClassImplementation;

import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.NonNull;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jetbrains.annotations.Contract;

import java.util.*;
import java.util.function.Function;
import java.util.function.Predicate;
import java.util.function.Consumer;

/**
 * <h2>XLSTableParser</h2>
 * <p>Класс-парсер для {@link XSSFWorkbook} из библиотеки <a href="https://poi.apache.org/">Apache POI</a></p>
 * <p>
 * Предоставляет интерфейс-алгоритм прохождения по рядам таблицы с четырьмя главными пунктами:
 * <ul>
 *     <li>Найти ряд, с которого начнётся чтение (через условие и, возможно, пропуск N рядов после его выполнения)</li>
 *     <li>Задать условие прекращения чтения</li>
 *     <li>Прочесть каждую колонку каждого ряда в DTO</li>
 *     <li>Упаковать каждый новый DTO в коллекцию или передать в лямбду-{@link Consumer}</li>
 * </ul>
 * </p>
 * После всего требует экземпляр класса {@link XSSFWorkbook}, из которого попытается прочесть поля.
 * <h3>Примеры кода:</h3>
 * <p><b>1)</b> Простейишй пример. Парсинг в словарь и вывод в консоль</p>
 * <pre>{@code XLSTableParser.fromSheet("Лист 1") // Листы также можно выбирать по их порядковому номеру
 *         .findRowThat(row -> row.getCell(row.getFirstCellNum()).getStringCellValue().contains("Заголовок таблицы"))
 *         .thenSkip(1) // Пропускаем до следующей после заголовка таблицы строки
 *         .endIfCell(0).isNull().or().isEmpty().or().isNotNumeric() // Для задания строки, на которой завершится парсинг.
 *                                                                   // Строка, удовлетворившая условиям и все после, парситься не будет
 *
 *         .getEntityFrom(() -> new LinkedHashMap<String, String>()) // Задаём, откуда взять или как создать DTO, куда будут закладываться данные
 *                                                                   // В этом случае, в качестве объекта взяли словарь строк
 *                                                                   // Supplier переданный в метод, создаёт новый объект для наполнения
 *
 *         .thenForColumnStringifiedValue((dto, s) -> dto.put(s, s)) // Для работы с данными существует несколько методов:
 *                                                                   // Можно работать напрямую с ячейкой, получая данные из неё
 *                                                                   // Можно запросить данные, приведённые к строке
 *         .thenForColumn((dto, cell) -> dto.put(cell.getRawValue(), Integer.toString(Double.valueOf(cell.getNumericCellValue()).intValue())))
 *
 *         .thenPutInto(System.out::println) // Задаётся Consumer, который будет принимать созданные в процессе ДТО
 *         .parse(xssfWorkbook) // С помощью написанной выше инструкции можно как распарсить книгу,
 *                              // так и запомнить в объект класса XLSTableParser, чтобы использовать
 *                              // на нескольких документах, если необходимо.}</pre>
 */
public class XLSTableParser {
    private final List<TableFiller<?>> fillers;

    /**
     * Первый метод для задания инструкции парсинга XLSX
     * @param sheetName Имя листа в XLSX
     * @return {@link StartCondition} Объект, задающий условия нахождения первого ряда
     */
    @Contract("_ -> new")
    public static @NonNull StartCondition fromSheet(@NonNull String sheetName) {
        return new StartCondition(new ArrayList<>(), (workbook) -> workbook.getSheet(sheetName));
    }

    /**
     * Первый метод для задания инструкции парсинга XLSX
     * @param sheetNumber Номер листа в XLSX
     * @return {@link StartCondition} Объект, задающий условия нахождения первого ряда
     */
    @Contract("_ -> new")
    public static @NonNull StartCondition fromSheet(int sheetNumber) {
        return new StartCondition(new ArrayList<>(), (workbook) -> workbook.getSheet(workbook.getSheetName(sheetNumber)));
    }

    private XLSTableParser(@NonNull List<TableFiller<?>> fillers) {
        this.fillers = fillers;
    }

    /**
     * Первый метод для задания инструкции парсинга XLSX. Нужен для задания инструкций для другого листа в XLSX книге
     * @param sheetName Имя следующего листа в XLSX
     * @return {@link StartCondition} Объект, задающий условия нахождения первого ряда
     */
    public StartCondition thenFromSheet(@NonNull String sheetName) {
        return new StartCondition(new LinkedList<>(fillers), (workbook) -> workbook.getSheet(sheetName));
    }

    /**
     * Первый метод для задания инструкции парсинга XLSX. Нужен для задания инструкций для другого листа в XLSX книге
     * @param sheetNumber Номер следующего листа в XLSX
     * @return {@link StartCondition} Объект, задающий условия нахождения первого ряда
     */
    public StartCondition thenFromSheet(int sheetNumber) {
        return new StartCondition(new LinkedList<>(fillers), (workbook) -> workbook.getSheet(workbook.getSheetName(sheetNumber)));
    }

    /**
     * Первый метод для задания инструкции парсинга XLSX. Нужен для задания инструкций для того же листа в XLSX книге
     * <b>с нулевого ряда</b>
     * @return {@link StartCondition} Объект, задающий условия нахождения первого ряда
     */
    public StartCondition thenRestartSameSheet() {
        return new StartCondition(new LinkedList<>(fillers), fillers.get(fillers.size() - 1).getSheetGetter());
    }

    /**
     * Первый метод для задания инструкции парсинга XLSX. Нужен для задания инструкций для того же листа в XLSX книге
     * <b>с ряда, где закончилась предыдущая таблица</b>
     * @return {@link StartCondition} Объект, задающий условия нахождения первого ряда
     */
    public StartCondition thenContinueSameSheet() {
        fillers.get(fillers.size() - 1).continueWhereEnded();
        return new StartCondition(new LinkedList<>(fillers), fillers.get(fillers.size() - 1).getSheetGetter());
    }

    /**
     * Метод, запускающий парсинг по вышезаданной инструкции.
     * @param book Книга, которая подвергнется парсингу по инструкции выше
     */
    public void parse(@NonNull XSSFWorkbook book) {
        int startFrom = 0;
        for (TableFiller<?> filler : fillers) {
            startFrom = filler.fillFrom(book, startFrom);
        }
    }

    /**
     * Класс, предоставляющий метод для нахождения первого ряда.
     */
    @AllArgsConstructor(access = AccessLevel.PRIVATE)
    public static final class StartCondition {
        private final List<TableFiller<?>> fillers;
        private final Function<XSSFWorkbook, XSSFSheet> sheetGetter;

        /**
         * Метод, принимающий условие взятия ряда (и всех последующих рядов) в работу.
         * @param startIf Предикат-условие, при котором этот и дальнейшие ряды будут приняты в обработку. Если этот и определённое
         *                количество последующих рядов является лишь маркером или заголовком, по которому была найдена таблица,
         *                их можно пропустить, указав их количество в следующем методе
         * @return {@link Skipper} Класс для пропуска лишних рядов
         */
        @Contract("_ -> new")
        public @NonNull Skipper findRowThat(@NonNull Predicate<XSSFRow> startIf) {
            return new Skipper(fillers, sheetGetter, startIf);
        }
    }

    /**
     * Класс, задающий количество рядов для пропуска.
     */
    @AllArgsConstructor(access = AccessLevel.PRIVATE)
    public static final class Skipper {
        private final List<TableFiller<?>> fillers;
        private final Function<XSSFWorkbook, XSSFSheet> sheetGetter;
        private final Predicate<XSSFRow> filter;

        @Contract("_ -> new")
        public @NonNull EndCondition thenSkip(int skip) {
            if (skip < 0) throw new IllegalArgumentException("Количество рядов для пропуска не может быть отрицательным, но пришло " + skip);
            return new EndCondition(fillers, sheetGetter, filter, skip);
        }
        public @NonNull EndCondition noSkip() {
            return thenSkip(0);
        }
    }
    @AllArgsConstructor(access = AccessLevel.PRIVATE)
    public static final class EndCondition {
        private final List<TableFiller<?>> fillers;
        private final Function<XSSFWorkbook, XSSFSheet> sheetGetter;
        private final Predicate<XSSFRow> filter;
        private final int skip;

        @Contract("_ -> new")
        public @NonNull EntityGetter endIf(@NonNull Predicate<XSSFRow> rowDecliner) {
            return new EntityGetter(fillers, sheetGetter, filter, skip, rowDecliner);
        }
        @Contract("_, _ -> new")
        public @NonNull EntityGetter endIfStringValueOfCell(int cellNum, @NonNull Predicate<String> rowDecliner) {
            return new EntityGetter(fillers, sheetGetter, filter, skip, row -> rowDecliner.test(new DataFormatter().formatCellValue(row.getCell(cellNum))));
        }
        @Contract("_ -> new")
        public @NonNull CellConditionBuilder endIfCell(int cellNum) {
            return new CellConditionBuilder(fillers, sheetGetter, filter, skip, cellNum, predicate->predicate);
        }
    }
    @AllArgsConstructor(access = AccessLevel.PRIVATE)
    public static final class CellConditionBuilder extends Condition<EndConditionLinker, CellConditionBuilder> {
        private final List<TableFiller<?>> fillers;
        private final Function<XSSFWorkbook, XSSFSheet> sheetGetter;
        private final Predicate<XSSFRow> filter;
        private final int skip;
        private final int cellNum;
        private final Function<Predicate<XSSFRow>, Predicate<XSSFRow>> initial;
        @Contract("_ -> new")
        protected @NonNull XLSTableParser.EndConditionLinker test(@NonNull Predicate<XSSFCell> condition) {
            return new EndConditionLinker(fillers, sheetGetter, filter, skip, cellNum, initial.apply(row -> condition.test(row.getCell(cellNum))));
        }
    }
    public static final class EndConditionLinker extends ConditionLinker<CellConditionBuilder, EndConditionLinker>{
        private final List<TableFiller<?>> fillers;
        private final Function<XSSFWorkbook, XSSFSheet> sheetGetter;
        private final Predicate<XSSFRow> filter;
        private final int skip;

        private EndConditionLinker(List<TableFiller<?>> fillers, Function<XSSFWorkbook, XSSFSheet> sheetGetter, Predicate<XSSFRow> filter, int skip, int cellNum, Predicate<XSSFRow> initial) {
            super(cellNum, initial);
            this.fillers = fillers;
            this.sheetGetter = sheetGetter;
            this.filter = filter;
            this.skip = skip;
        }


        @Override
        protected CellConditionBuilder goBack(int cellNum, Function<Predicate<XSSFRow>, Predicate<XSSFRow>> transformer) {
            return new CellConditionBuilder(fillers, sheetGetter,filter,skip,cellNum,transformer);
        }
        @Contract("_ -> new")
        public <T> @NonNull EntityFiller<T> getEntityFrom(@NonNull UncheckedSupplier<T> generator) {
            return new EntityFiller<>(fillers, sheetGetter, filter, skip, initial, generator);
        }
    }
    @AllArgsConstructor(access = AccessLevel.PRIVATE)
    public static final class EntityGetter {
        private final List<TableFiller<?>> fillers;
        private final Function<XSSFWorkbook, XSSFSheet> sheetGetter;
        private final Predicate<XSSFRow> filter;
        private final int skip;
        private final Predicate<XSSFRow> rowDecliner;

        @Contract("_ -> new")
        public <T> @NonNull EntityFiller<T> getEntityFrom(@NonNull UncheckedSupplier<T> generator) {
            return new EntityFiller<>(fillers, sheetGetter, filter, skip, rowDecliner, generator);
        }
    }
    @AllArgsConstructor(access = AccessLevel.PRIVATE)
    public static final class EntityFiller<T> {
        private final List<TableFiller<?>> fillers;
        private final Function<XSSFWorkbook, XSSFSheet> sheetGetter;
        private final Predicate<XSSFRow> filter;
        private final int skip;
        private final Predicate<XSSFRow> rowDecliner;
        private final UncheckedSupplier<T> generator;
        private final List<UncheckedBiConsumer<T, XSSFCell>> columnFillers = new ArrayList<>();

        public EntityFiller<T> thenForColumn(@NonNull UncheckedBiConsumer<T, XSSFCell> filler) {
            columnFillers.add(filler);
            return this;
        }
        public EntityFiller<T> thenForColumnStringifiedValue(@NonNull UncheckedBiConsumer<T, String> filler) {
            columnFillers.add((dto, cell) -> filler.acceptUnchecked(dto, new DataFormatter().formatCellValue(cell)));
            return this;
        }
        @Contract("_ -> new")
        public @NonNull XLSTableParser thenPutInto(@NonNull UncheckedConsumer<T> consumer) {
            fillers.add(new TableFiller<>(sheetGetter, columnFillers, generator, filter, rowDecliner, skip, consumer));
            return new XLSTableParser(Collections.unmodifiableList(fillers));
        }
        @Contract("_ -> new")
        public @NonNull XLSTableParser thenPutInto(@NonNull Collection<? super T> collection) {
            fillers.add(new TableFiller<>(sheetGetter, columnFillers, generator, filter, rowDecliner, skip, collection::add));
            return new XLSTableParser(Collections.unmodifiableList(fillers));
        }
    }
    private static abstract class Condition<LINKER extends ConditionLinker<CONDITION, LINKER>, CONDITION extends Condition<LINKER, CONDITION>> {
        private static final Predicate<XSSFCell> IS_NUMERIC = cell -> cell.getCellType().equals(CellType.NUMERIC);
        private static final Predicate<XSSFCell> IS_STRING = cell -> cell.getCellType().equals(CellType.STRING);
        private static final Predicate<XSSFCell> IS_BLANK = cell -> cell.getCellType().equals(CellType.BLANK);
        private static final Predicate<XSSFCell> IS_BOOLEAN = cell -> cell.getCellType().equals(CellType.BOOLEAN);
        private static final Predicate<XSSFCell> IS_ERROR = cell -> cell.getCellType().equals(CellType.ERROR);
        private static final Predicate<XSSFCell> IS_FORMULA = cell -> cell.getCellType().equals(CellType.FORMULA);
        private static final Predicate<XSSFCell> IS_NULL = Objects::isNull;
        protected abstract LINKER test(Predicate<XSSFCell> condition);
        @Contract(" -> new")
        public @NonNull LINKER isNumeric() {
            return test(IS_NUMERIC);
        }
        @Contract(" -> new")
        public @NonNull LINKER isNotNumeric() {
            return test(IS_NUMERIC.negate());
        }
        @Contract(" -> new")
        public @NonNull LINKER isString() {
            return test(IS_STRING);
        }
        @Contract(" -> new")
        public @NonNull LINKER isNotString() {
            return test(IS_STRING.negate());
        }
        @Contract(" -> new")
        public @NonNull LINKER isEmpty() {
            return test(IS_BLANK);
        }
        @Contract(" -> new")
        public @NonNull LINKER isNotEmpty() {
            return test(IS_BLANK.negate());
        }
        @Contract(" -> new")
        public @NonNull LINKER isBoolean() {
            return test(IS_BOOLEAN);
        }
        @Contract(" -> new")
        public @NonNull LINKER isNotBoolean() {
            return test(IS_BOOLEAN.negate());
        }
        @Contract(" -> new")
        public @NonNull LINKER isFormula() {
            return test(IS_FORMULA);
        }

        @Contract(" -> new")
        public @NonNull LINKER isNotFormula() {
            return test(IS_FORMULA.negate());
        }
        @Contract(" -> new")
        public @NonNull LINKER isError() {
            return test(IS_ERROR);
        }
        @Contract(" -> new")
        public @NonNull LINKER isNotError() {
            return test(IS_ERROR.negate());
        }
        @Contract(" -> new")
        public @NonNull LINKER isNull() {
            return test(IS_NULL);
        }
        @Contract(" -> new")
        public @NonNull LINKER isNotNull() {
            return test(IS_NULL.negate());
        }
        @Contract("_ -> new")
        public @NonNull LINKER stringValueEquals(@NonNull String other) {
            return test(IS_STRING.and(cell -> cell.getStringCellValue().equals(other)));
        }
        @Contract("_ -> new")
        public @NonNull LINKER stringValueEqualsIgnoreCase(@NonNull String other) {
            return test(IS_STRING.and(cell -> cell.getStringCellValue().equalsIgnoreCase(other)));
        }
        @Contract("_ -> new")
        public @NonNull LINKER stringValueContains(@NonNull String other) {
            return test(IS_STRING.and(cell -> cell.getStringCellValue().contains(other)));
        }
        @Contract("_ -> new")
        public @NonNull LINKER stringValueContainsIgnoreCase(@NonNull String other) {
            return test(IS_STRING.and(cell -> cell.getStringCellValue().toLowerCase().contains(other.toLowerCase())));
        }
    }
    private static abstract class ConditionLinker<CONDITION extends Condition<LINKER, CONDITION>, LINKER extends ConditionLinker<CONDITION, LINKER>> {
        private final int cellNum;
        protected final Predicate<XSSFRow> initial;

        private ConditionLinker(int cellNum, Predicate<XSSFRow> initial) {
            this.cellNum = cellNum;
            this.initial = initial;
        }

        protected abstract CONDITION goBack(int cellNum, Function<Predicate<XSSFRow>, Predicate<XSSFRow>> transformer);

        public @NonNull CONDITION and() {
            return goBack(cellNum, initial::and);
        }
        public @NonNull CONDITION andOtherCell(int otherCellNum) {
            return goBack(otherCellNum, initial::and);
        }
        public @NonNull CONDITION or() {
            return goBack(cellNum, initial::or);
        }
        public @NonNull CONDITION orOtherCell(int otherCellNum) {
            return goBack(otherCellNum, initial::or);
        }
    }
}