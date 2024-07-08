package com.kanayaya.XLSParse.InnerClassImplementation;

import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.NonNull;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jetbrains.annotations.Contract;
import org.jetbrains.annotations.NotNull;

import java.util.*;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Predicate;

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
 * <p><b>1)</b> Простейший пример. Парсинг в словарь и вывод в консоль</p>
 * <pre>{@code XLSTableParser.fromSheet("Лист 1") // Листы также можно выбирать по их порядковому номеру
 *         .findRowThat(row -> row.getCell(row.getFirstCellNum())
 *             .getStringCellValue().contains("Заголовок таблицы"))
 *         .thenSkip(1) // Пропускаем до следующей после заголовка таблицы строки
 *
 *         // Для задания строки, на которой завершится парсинг.
 *         // Строка, удовлетворившая условиям и все после, парситься не будет
 *         .endIfCell(0).isNull().or().isEmpty().or().isNotNumeric()
 *
 *         // Задаём, откуда взять или как создать DTO, куда будут закладываться данные
 *         // В этом случае, в качестве объекта взяли словарь строк
 *         // Supplier переданный в метод, создаёт новый объект для наполнения
 *         .getEntityFrom(() -> new LinkedHashMap<String, String>())
 *
 *         // Для работы с данными существует несколько методов:
 *         // Можно работать напрямую с ячейкой, получая данные из неё
 *         // Можно запросить данные, приведённые к строке
 *         .thenForColumnStringifiedValue((dto, s) -> dto.put(s, s))
 *         .thenForColumn((dto, cell) -> dto.put(cell.getRawValue(),
 *             Integer.toString(
 *                 Double.valueOf(cell.getNumericCellValue()).intValue())))
 *
 *         // Задаётся Consumer, который будет принимать созданные в процессе ДТО
 *         .thenPutInto(System.out::println)
 *
 *         // С помощью написанной выше инструкции можно как распарсить книгу,
 *         // так и запомнить в объект класса XLSTableParser, чтобы использовать
 *         // на нескольких документах, если необходимо.
 *         .parse(xssfWorkbook)}</pre>
 */
@Slf4j
public class XLSTableParser {
    /**
     * Когда задана инструкция, сюда кладётся её формальное объявление как экземпляр класса {@link TableFiller}
     * <p>Таких инструкций может быть несколько -- они выполнятся последовательно</p>
     */
    private final TableFiller<?> lastFiller;
    private final TransitiveBiFunction<XSSFWorkbook, Integer, Integer> parserChain;
    private final Function<XSSFWorkbook, XSSFSheet> sheetGetter;

    /**
     * Первый метод для задания инструкции парсинга XLSX
     * @param sheetName Имя листа в XLSX
     * @return {@link StartConditionGetter} Объект, задающий условия нахождения первого ряда
     */
    @Contract("_ -> new")
    public static @NonNull StartConditionGetter fromSheet(@NonNull String sheetName) {
        return new StartConditionGetter((workbook, i) -> i, (workbook) -> workbook.getSheet(sheetName));
    }

    /**
     * Первый метод для задания инструкции парсинга XLSX
     * @param sheetNumber Номер листа в XLSX
     * @return {@link StartConditionGetter} Объект, задающий условия нахождения первого ряда
     */
    @Contract("_ -> new")
    public static @NonNull StartConditionGetter fromSheet(int sheetNumber) {
        return new StartConditionGetter((workbook, i) -> i, (workbook) -> workbook.getSheet(workbook.getSheetName(sheetNumber)));
    }

    private XLSTableParser(TableFiller<?> lastFiller, TransitiveBiFunction<XSSFWorkbook, Integer, Integer> parser, Function<XSSFWorkbook, XSSFSheet> sheetGetter) {
        this.lastFiller = lastFiller;
        this.parserChain = parser;
        this.sheetGetter = sheetGetter;
    }

    /**
     * Метод для задания инструкции парсинга следующей таблицы (или той же) из XLSX. Нужен для задания инструкций для другого листа в XLSX книге
     * @param sheetName Имя следующего листа в XLSX
     * @return {@link StartConditionGetter} Объект, задающий условия нахождения первого ряда
     */
    public @NonNull StartConditionGetter thenFromSheet(@NonNull String sheetName) {
        return new StartConditionGetter(parserChain.andThen(lastFiller::fillFrom), (workbook) -> workbook.getSheet(sheetName));
    }

    /**
     * Метод для задания инструкции парсинга следующей таблицы (или той же) из XLSX. Нужен для задания инструкций для другого листа в XLSX книге
     * @param sheetNumber Номер следующего листа в XLSX
     * @return {@link StartConditionGetter} Объект, задающий условия нахождения первого ряда
     */
    public @NonNull StartConditionGetter thenFromSheet(int sheetNumber) {
        return new StartConditionGetter(parserChain.andThen(lastFiller::fillFrom), (workbook) -> workbook.getSheet(workbook.getSheetName(sheetNumber)));
    }

    /**
     * Метод для задания инструкции парсинга следующей таблицы (или той же) из XLSX. Нужен для задания инструкций для того же листа в XLSX книге
     * <b>с нулевого ряда</b>
     * @return {@link StartConditionGetter} Объект, задающий условия нахождения первого ряда
     */
    public @NonNull StartConditionGetter thenRestartSameSheet() {
        return new StartConditionGetter(parserChain.andThen(lastFiller::fillFrom), sheetGetter);
    }

    /**
     * Метод для задания инструкции парсинга следующей таблицы из XLSX. Нужен для задания инструкций для того же листа в XLSX книге
     * <b>с ряда, где закончилась предыдущая таблица</b>
     * @return {@link StartConditionGetter} Объект, задающий условия нахождения первого ряда
     */
    public @NonNull StartConditionGetter thenContinueSameSheet() {
        return new StartConditionGetter(parserChain.andThen(lastFiller::fillContinuing), sheetGetter);
    }

    /**
     * Метод, запускающий парсинг по инструкции, заданной до того, как прийти к этому методу.
     * @param book Книга, которая подвергнется парсингу по заданной инструкции
     */
    public void parse(@NonNull XSSFWorkbook book) {
        parserChain.andThen(lastFiller::fillFrom).apply(book, 0);
    }

    /**
     * Класс, предоставляющий метод для нахождения первого ряда.
     */
    @AllArgsConstructor(access = AccessLevel.PRIVATE)
    public static final class StartConditionGetter {
        private final TransitiveBiFunction<XSSFWorkbook, Integer, Integer> parser;
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
            return new Skipper(parser, sheetGetter, startIf);
        }

        /**
         * @param cellNum Номер столбца <b>ИЛИ</b> код из класса {@link CellCodes} ({@link CellCodes#FIRST} или {@link CellCodes#LAST})
         * @return {@link StartCondition} Класс для задания условия нахождения ряда по столбцам
         * @throws IllegalArgumentException В случае отрицательного номера столбца, не соответствующего коду из класса {@link CellCodes}
         */
        @Contract("_ -> new")
        public @NonNull StartCondition findRowWhereCell(int cellNum) {
            if (cellNum < -2) throw new IllegalArgumentException("Неверный номер столбца: " + cellNum);
            return new StartCondition(parser, sheetGetter, cellNum, rowCondition->rowCondition);
        }
    }

    /**
     * Класс, описывающий условие начала парсинга
     */
    @AllArgsConstructor(access = AccessLevel.PRIVATE)
    public static final class StartCondition extends Condition<StartConditionLinker, StartCondition> {
        private final TransitiveBiFunction<XSSFWorkbook, Integer, Integer> parser;
        private final Function<XSSFWorkbook, XSSFSheet> sheetGetter;
        private final int cellNum;
        private final Function<Predicate<XSSFRow>, Predicate<XSSFRow>> initial;

        @Override
        protected @NonNull StartConditionLinker test(@NotNull Predicate<XSSFCell> condition) {
            return new StartConditionLinker(parser, sheetGetter, cellNum,
                    initial.apply(row -> condition.test(row.getCell(
                            cellNum == CellCodes.FIRST?
                                    row.getFirstCellNum() :
                                    cellNum == CellCodes.LAST?
                                            row.getLastCellNum() :
                                            cellNum))));
        }
    }

    /**
     * Класс, описывающий связку нескольких условий и переход далее по алгоритму
     */
    public static final class StartConditionLinker extends ConditionLinker<StartCondition, StartConditionLinker> {
        private final TransitiveBiFunction<XSSFWorkbook, Integer, Integer> parser;
        private final Function<XSSFWorkbook, XSSFSheet> sheetGetter;
        private StartConditionLinker(@NonNull TransitiveBiFunction<XSSFWorkbook, Integer, Integer> parser, @NonNull Function<XSSFWorkbook, XSSFSheet> sheetGetter, int cellNum, @NonNull Predicate<XSSFRow> initial) {
            super(cellNum, initial);
            this.parser = parser;
            this.sheetGetter = sheetGetter;
        }

        @Override
        protected @NonNull StartCondition goBack(int cellNum, @NotNull Function<Predicate<XSSFRow>, Predicate<XSSFRow>> transformer) {
            return new StartCondition(parser, sheetGetter, cellNum, transformer);
        }

        /**
         * @param skip Сколько строк таблицы пропустить после нахождения ряда по условию
         * @return Класс, принимающий условие окончания парсинга
         */
        @Contract("_ -> new")
        public @NonNull EndConditionGetter thenSkip(int skip) {
            if (skip < 0) throw new IllegalArgumentException("Количество рядов для пропуска не может быть отрицательным, но пришло " + skip);
            return new EndConditionGetter(parser, sheetGetter, initial, skip);
        }
        /**
         * Метод начала обработки без пропуска рядов
         * @return Класс, принимающий условие окончания парсинга
         */
        public @NonNull EndConditionGetter noSkip() {
            return thenSkip(0);
        }
    }

    /**
     * Класс, задающий количество рядов для пропуска.
     */
    @AllArgsConstructor(access = AccessLevel.PRIVATE)
    public static final class Skipper {
        private final TransitiveBiFunction<XSSFWorkbook, Integer, Integer> parser;
        private final Function<XSSFWorkbook, XSSFSheet> sheetGetter;
        private final Predicate<XSSFRow> filter;

        /**
         * @param skip Сколько строк таблицы пропустить после нахождения ряда по условию
         * @return Класс, принимающий условие окончания парсинга
         */
        @Contract("_ -> new")
        public @NonNull EndConditionGetter thenSkip(int skip) {
            if (skip < 0) throw new IllegalArgumentException("Количество рядов для пропуска не может быть отрицательным, но пришло " + skip);
            return new EndConditionGetter(parser, sheetGetter, filter, skip);
        }
        /**
         * Метод начала обработки без пропуска рядов
         * @return Класс, принимающий условие окончания парсинга
         */
        public @NonNull EndConditionGetter noSkip() {
            return thenSkip(0);
        }
    }
    @AllArgsConstructor(access = AccessLevel.PRIVATE)
    public static final class EndConditionGetter {
        private final TransitiveBiFunction<XSSFWorkbook, Integer, Integer> parser;
        private final Function<XSSFWorkbook, XSSFSheet> sheetGetter;
        private final Predicate<XSSFRow> filter;
        private final int skip;

        /**
         * @param rowDecliner Предикат-условие окончания парсинга
         * @return {@link EntityGetter} Класс для выставления типа DTO и метода его создания
         */
        @Contract("_ -> new")
        public @NonNull EntityGetter endIf(@NonNull Predicate<XSSFRow> rowDecliner) {
            return new EntityGetter(parser, sheetGetter, filter, skip, rowDecliner);
        }

        /**
         * @param cellNum Номер столбца <b>ИЛИ</b> код из класса {@link CellCodes} ({@link CellCodes#FIRST} или {@link CellCodes#LAST})
         * @return {@link StartCondition} Класс для задания условия нахождения ряда по столбцам
         * @throws IllegalArgumentException В случае отрицательного номера столбца, не соответствующего коду из класса {@link CellCodes}
         */
        @Contract("_ -> new")
        public @NonNull EndCondition endIfCell(int cellNum) {
            if (cellNum < -2) throw new IllegalArgumentException("Неверный номер столбца: " + cellNum);
            return new EndCondition(parser, sheetGetter, filter, skip, cellNum, predicate->predicate);
        }
    }
    @AllArgsConstructor(access = AccessLevel.PRIVATE)
    public static final class EndCondition extends Condition<EndConditionLinker, EndCondition> {
        private final TransitiveBiFunction<XSSFWorkbook, Integer, Integer> parser;
        private final Function<XSSFWorkbook, XSSFSheet> sheetGetter;
        private final Predicate<XSSFRow> filter;
        private final int skip;
        private final int cellNum;
        private final Function<Predicate<XSSFRow>, Predicate<XSSFRow>> initial;
        @Contract("_ -> new")
        protected @NonNull EndConditionLinker test(@NonNull Predicate<XSSFCell> condition) {
            return new EndConditionLinker(parser, sheetGetter, filter, skip, cellNum,
                    initial.apply(row -> condition.test(row.getCell(
                            cellNum == CellCodes.FIRST?
                                    row.getFirstCellNum() :
                                    cellNum == CellCodes.LAST?
                                            row.getLastCellNum() :
                                            cellNum))));
        }
    }
    public static final class EndConditionLinker extends ConditionLinker<EndCondition, EndConditionLinker>{
        private final TransitiveBiFunction<XSSFWorkbook, Integer, Integer> parser;
        private final Function<XSSFWorkbook, XSSFSheet> sheetGetter;
        private final Predicate<XSSFRow> filter;
        private final int skip;

        private EndConditionLinker(@NonNull TransitiveBiFunction<XSSFWorkbook, Integer, Integer> parser, Function<XSSFWorkbook, XSSFSheet> sheetGetter, @NonNull Predicate<XSSFRow> filter, int skip, int cellNum, @NonNull Predicate<XSSFRow> initial) {
            super(cellNum, initial);
            this.parser = parser;
            this.sheetGetter = sheetGetter;
            this.filter = filter;
            this.skip = skip;
        }
        @Override
        protected @NonNull EndCondition goBack(int cellNum, @NotNull Function<Predicate<XSSFRow>, @NonNull Predicate<XSSFRow>> transformer) {
            return new EndCondition(parser, sheetGetter,filter,skip,cellNum,transformer);
        }
        /**
         * @param generator {@link UncheckedSupplier} Генератор DTO
         * @param <T> Тип DTO
         * @return {@link EntityFillerSequential} Класс для заполнения DTO заданного типа
         */
        @Contract("_ -> new")
        public <T> @NonNull EntityFillerVariant<T> getEntityFrom(@NonNull UncheckedSupplier<T> generator) {
            return new EntityFillerVariant<>(parser, sheetGetter, filter, skip, initial, generator);
        }
    }
    @AllArgsConstructor(access = AccessLevel.PRIVATE)
    public static final class EntityGetter {
        private final TransitiveBiFunction<XSSFWorkbook, Integer, Integer> parser;
        private final Function<XSSFWorkbook, XSSFSheet> sheetGetter;
        private final Predicate<XSSFRow> filter;
        private final int skip;
        private final Predicate<XSSFRow> rowDecliner;

        /**
         * @param generator {@link UncheckedSupplier} Генератор DTO
         * @param <T> Тип DTO
         * @return {@link EntityFillerSequential} Класс для заполнения DTO заданного типа
         */
        @Contract("_ -> new")
        public <T> @NonNull EntityFillerVariant<T> getEntityFrom(@NonNull UncheckedSupplier<T> generator) {
            return new EntityFillerVariant<>(parser, sheetGetter, filter, skip, rowDecliner, generator);
        }
    }

    @AllArgsConstructor(access = AccessLevel.PRIVATE)
    public static final class EntityFillerVariant<T> {
        private final TransitiveBiFunction<XSSFWorkbook, Integer, Integer> parser;
        private final Function<XSSFWorkbook, XSSFSheet> sheetGetter;
        private final Predicate<XSSFRow> filter;
        private final int skip;
        private final Predicate<XSSFRow> rowDecliner;
        private final UncheckedSupplier<T> generator;
        public EntityFillerSequential<T> thenForNextColumn(@NonNull UncheckedBiConsumer<T, XSSFCell> filler) {
            UncheckedBiConsumer<T, XSSFRow> columnFiller = (dto, row) -> filler.acceptUnchecked(dto, row.getCell(row.getFirstCellNum()));
            return new EntityFillerSequential<>(parser, sheetGetter, filter, skip, rowDecliner, generator, columnFiller, 1);
        }
        public EntityFillerSequential<T> thenForNextColumnStringified(@NonNull UncheckedBiConsumer<T, String> filler) {
            return thenForNextColumn((dto, cell) -> filler.acceptUnchecked(dto, new DataFormatter().formatCellValue(cell)));
        }
        public EntityFillerNumberChooser<T> thenForColumn(int cellNum, @NonNull UncheckedBiConsumer<T, XSSFCell> filler) {
            UncheckedBiConsumer<T, XSSFRow> columnFiller = (dto, row) -> {
                XSSFCell cell = row.getCell(cellNum == CellCodes.FIRST? row.getFirstCellNum() : cellNum == CellCodes.LAST? row.getLastCellNum() : cellNum);
                if (cell == null) log.warn(String.format("Столбец ряда %d не содержит ячейку %d (null)", row.getRowNum(), cellNum));
                filler.acceptUnchecked(dto, cell);
            };
            return new EntityFillerNumberChooser<>(parser, sheetGetter, filter, skip, rowDecliner, generator, columnFiller);
        }

        /**
         * Принимает номер столбца и {@link UncheckedBiConsumer} для заполнения DTO его строковым значением
         * @param cellNum Номер столбца
         * @param filler {@link UncheckedBiConsumer} для заполнения DTO
         * @return Такой же {@link EntityFillerNumberChooser} для дальнейшего заполнения DTO
         * или перехода к следующему этапу инструкции
         */
        @Contract("_, _ -> new")
        public @NotNull EntityFillerNumberChooser<T> thenForColumnStringified(int cellNum, @NonNull UncheckedBiConsumer<T, String> filler) {
            return thenForColumn(cellNum, (dto, cell) -> filler.acceptUnchecked(dto, new DataFormatter().formatCellValue(cell)));
        }
    }

    /**
     * Класс описывает наполнитель для прохождения по столбцам таблицы <br>
     * по их номерам и наполнения с их помощью DTO
     *
     * @param <T> Тип DTO
     */
    public static final class EntityFillerNumberChooser<T> extends EntityFiller<T> {
        private EntityFillerNumberChooser(TransitiveBiFunction<XSSFWorkbook, Integer, Integer> parser, Function<XSSFWorkbook, XSSFSheet> sheetGetter, Predicate<XSSFRow> filter, int skip, Predicate<XSSFRow> rowDecliner, UncheckedSupplier<T> generator, UncheckedBiConsumer<T, XSSFRow> columnFiller) {
            super(parser, sheetGetter, filter, skip, rowDecliner, generator, columnFiller);
        }

        /**
         * Метод для внесения способа заполнения DTO из ячейки.
         * @param cellNum Номер столбца ряда начиная с 0. Или используйте {@link CellCodes#FIRST} или {@link CellCodes#LAST} для первого и последнего столбца соответственно
         * @param filler Лямбда, говорящая о том, как положить содержимое ячейки в DTO
         * @return Себя же, для дальнейшего заполнения
         */
        @Contract("_, _ -> new")
        public @NotNull EntityFillerNumberChooser<T> thenForColumn(int cellNum, @NonNull UncheckedBiConsumer<T, XSSFCell> filler) {
            UncheckedBiConsumer<T, XSSFRow> newFiller = columnFiller.andThen((dto, row) -> {
                XSSFCell cell = row.getCell(cellNum == CellCodes.FIRST? row.getFirstCellNum() : cellNum == CellCodes.LAST? row.getLastCellNum() : cellNum);
                if (cell == null) log.warn(String.format("Столбец ряда %d не содержит ячейку %d (null)", row.getRowNum(), cellNum));
                filler.accept(dto, cell);
            });
            return new EntityFillerNumberChooser<>(parser, sheetGetter, filter, skip, rowDecliner, generator, newFiller);
        }
        /**
         * Метод для внесения способа заполнения DTO из строкового представления ячейки.
         * @param cellNum Номер столбца ряда начиная с 0. Или используйте {@link CellCodes#FIRST} или {@link CellCodes#LAST} для первого и последнего столбца соответственно
         * @param filler Лямбда, говорящая о том, как положить содержимое ячейки в DTO
         * @return Себя же, для дальнейшего заполнения
         */
        @Contract("_, _ -> new")
        public @NotNull EntityFillerNumberChooser<T> thenForColumnStringified(int cellNum, @NonNull UncheckedBiConsumer<T, String> filler) {
            return thenForColumn(cellNum, (dto, cell) -> filler.acceptUnchecked(dto, new DataFormatter().formatCellValue(cell)));
        }
    }

    /**
     * Класс описывает наполнитель для последовательного прохождения <br>
     * по столбцам таблицы и наполнения с их помощью DTO
     *
     * @param <T> Тип DTO
     */
    public static final class EntityFillerSequential<T> extends EntityFiller<T> {
        private final int cellNum;
        private EntityFillerSequential(TransitiveBiFunction<XSSFWorkbook, Integer, Integer> parser, Function<XSSFWorkbook, XSSFSheet> sheetGetter, Predicate<XSSFRow> filter, int skip, Predicate<XSSFRow> rowDecliner, UncheckedSupplier<T> generator, UncheckedBiConsumer<T, XSSFRow> columnFiller, int cellNum) {
            super(parser, sheetGetter, filter, skip, rowDecliner, generator, columnFiller);
            this.cellNum = cellNum;
        }
        /**
         * Метод для внесения способа заполнения DTO из ячейки.
         * @param filler Лямбда, говорящая о том, как положить содержимое ячейки в DTO
         * @return Себя же, для дальнейшего заполнения
         */
        @Contract("_ -> new")
        public @NotNull EntityFillerSequential<T> thenForNextColumn(@NonNull UncheckedBiConsumer<T, XSSFCell> filler) {
            UncheckedBiConsumer<T, XSSFRow> newFiller = columnFiller.andThen((dto, row) -> {
                int cellNum = row.getFirstCellNum() + this.cellNum;
                XSSFCell cell = row.getCell(cellNum);
                if (cell == null) log.warn(String.format("Столбец ряда %d не содержит ячейку %d (null)", row.getRowNum(), cellNum));
                filler.accept(dto, cell);
            });
            return new EntityFillerSequential<>(parser, sheetGetter, filter, skip, rowDecliner, generator, newFiller, cellNum + 1);
        }
        /**
         * Метод для внесения способа заполнения DTO из строкового представления ячейки.
         * @param filler Лямбда, говорящая о том, как положить содержимое ячейки в DTO
         * @return Себя же, для дальнейшего заполнения
         */
        @Contract("_ -> new")
        public @NotNull EntityFillerSequential<T> thenForNextColumnStringified(@NonNull UncheckedBiConsumer<T, String> filler) {
            return thenForNextColumn((dto, cell) -> filler.acceptUnchecked(dto, new DataFormatter().formatCellValue(cell)));
        }
    }

    /**
     * Класс описывает общую часть и состав всех классов-наполнителей для заполнения DTO.
     * <br>
     * Содержит поля и методы перехода к следующему шагу
     * @param <T> Тип DTO
     */
    @AllArgsConstructor(access = AccessLevel.PRIVATE)
    private static class EntityFiller<T> {
        protected final TransitiveBiFunction<XSSFWorkbook, Integer, Integer> parser;
        protected final Function<XSSFWorkbook, XSSFSheet> sheetGetter;
        protected final Predicate<XSSFRow> filter;
        protected final int skip;
        protected final Predicate<XSSFRow> rowDecliner;
        protected final UncheckedSupplier<T> generator;
        protected final UncheckedBiConsumer<T, XSSFRow> columnFiller;
        /**
         * Метод завершает набор условий парсинга и возвращает развилку выбора на новый цикл или начала парсинга
         * @param consumer Лямбда-потребитель для DTO созданного из каждого ряда
         * @return Развилка для задания следующей таблицы на парсинг или начала парсинга
         */
        @Contract("_ -> new")
        public @NonNull XLSTableParser thenPutInto(@NonNull UncheckedConsumer<? super T> consumer) {
            TableFiller<T> filler = new TableFiller<>(sheetGetter, columnFiller, generator, filter, rowDecliner, skip, consumer);
            return new XLSTableParser(filler, parser, sheetGetter);
        }
        /**
         * Метод завершает набор условий парсинга и возвращает развилку выбора на новый цикл или начала парсинга
         * @param collection Коллекция, в которую можно поместить DTO
         * @return Развилка для задания следующей таблицы на парсинг или начала парсинга
         */
        @Contract("_ -> new")
        public @NonNull XLSTableParser thenPutInto(@NonNull Collection<? super T> collection) {
            TableFiller<T> filler = new TableFiller<>(sheetGetter, columnFiller, generator, filter, rowDecliner, skip, collection::add);
            return new XLSTableParser(filler, parser, sheetGetter);
        }
    }

    /**
     * Класс описывает основную часть условия, применяемого к столбцу таблицы.
     * Содержит методы-проверки типа столбца и его значения.
     * @param <LINKER> Циклический дженерик для указания связи с парой-линкером
     * @param <CONDITION> Дженерик, в котором следует указать наследующий класс для выставления цикла
     */
    private static abstract class Condition<
            LINKER extends ConditionLinker<CONDITION, LINKER>,
            CONDITION extends Condition<LINKER, CONDITION>> {
        private static final Predicate<XSSFCell> IS_NUMERIC = cell -> cell.getCellType().equals(CellType.NUMERIC);
        private static final Predicate<XSSFCell> IS_STRING = cell -> cell.getCellType().equals(CellType.STRING);
        private static final Predicate<XSSFCell> IS_BLANK = cell -> cell.getCellType().equals(CellType.BLANK);
        private static final Predicate<XSSFCell> IS_BOOLEAN = cell -> cell.getCellType().equals(CellType.BOOLEAN);
        private static final Predicate<XSSFCell> IS_ERROR = cell -> cell.getCellType().equals(CellType.ERROR);
        private static final Predicate<XSSFCell> IS_FORMULA = cell -> cell.getCellType().equals(CellType.FORMULA);
        private static final Predicate<XSSFCell> IS_NULL = Objects::isNull;

        /**
         * Добавляет в инструкцию для парсера проверку условия. Само условие предоставляется аргументом.
         * @param condition Условие для проверки
         * @return Класс-линкер для указания связи со следующим условием или перехода далее по алгоритму
         */
        protected abstract LINKER test(@NonNull Predicate<XSSFCell> condition);

        /**
         * Добавляет в инструкцию для парсера проверку ранее указанного столбца на тип {@link CellType#NUMERIC}
         * <p>Проверяет что столбец содержит числовые данные</p>
         * @return Класс-линкер для указания связи со следующим условием или перехода далее по алгоритму
         */
        @Contract(" -> new")
        public @NonNull LINKER isNumeric() {
            return test(IS_NUMERIC);
        }
        /**
         * Добавляет в инструкцию для парсера проверку ранее указанного столбца на несоответствие типу {@link CellType#NUMERIC}
         * <p>Проверяет что столбец содержит не числовые данные</p>
         * @return Класс-линкер для указания связи со следующим условием или перехода далее по алгоритму
         */
        @Contract(" -> new")
        public @NonNull LINKER isNotNumeric() {
            return test(IS_NUMERIC.negate());
        }
        /**
         * Добавляет в инструкцию для парсера проверку ранее указанного столбца на тип {@link CellType#STRING}
         * <p>Проверяет что столбец содержит строчные данные</p>
         * @return Класс-линкер для указания связи со следующим условием или перехода далее по алгоритму
         */
        @Contract(" -> new")
        public @NonNull LINKER isString() {
            return test(IS_STRING);
        }
        /**
         * Добавляет в инструкцию для парсера проверку ранее указанного столбца на несоответствие типу {@link CellType#STRING}
         * <p>Проверяет что столбец содержит не строчные данные</p>
         * @return Класс-линкер для указания связи со следующим условием или перехода далее по алгоритму
         */
        @Contract(" -> new")
        public @NonNull LINKER isNotString() {
            return test(IS_STRING.negate());
        }
        /**
         * Добавляет в инструкцию для парсера проверку ранее указанного столбца на тип {@link CellType#BLANK}
         * <p>Проверяет что столбец <b>пуст</b></p>
         * @return Класс-линкер для указания связи со следующим условием или перехода далее по алгоритму
         */
        @Contract(" -> new")
        public @NonNull LINKER isEmpty() {
            return test(IS_BLANK);
        }
        /**
         * Добавляет в инструкцию для парсера проверку ранее указанного столбца на несоответствие типу {@link CellType#BLANK}
         * <p>Проверяет что столбец <b>что-то содержит</b></p>
         * @return Класс-линкер для указания связи со следующим условием или перехода далее по алгоритму
         */
        @Contract(" -> new")
        public @NonNull LINKER isNotEmpty() {
            return test(IS_BLANK.negate());
        }
        /**
         * Добавляет в инструкцию для парсера проверку ранее указанного столбца на тип {@link CellType#BOOLEAN}
         * <p>Проверяет что столбец содержит ИСТИНА или ЛОЖЬ</p>
         * @return Класс-линкер для указания связи со следующим условием или перехода далее по алгоритму
         */
        @Contract(" -> new")
        public @NonNull LINKER isBoolean() {
            return test(IS_BOOLEAN);
        }
        /**
         * Добавляет в инструкцию для парсера проверку ранее указанного столбца на несоответствие типу {@link CellType#BOOLEAN}
         * <p>Проверяет что столбец содержит любые данные кроме ИСТИНА и ЛОЖЬ</p>
         * @return Класс-линкер для указания связи со следующим условием или перехода далее по алгоритму
         */
        @Contract(" -> new")
        public @NonNull LINKER isNotBoolean() {
            return test(IS_BOOLEAN.negate());
        }
        /**
         * Добавляет в инструкцию для парсера проверку ранее указанного столбца на тип {@link CellType#FORMULA}
         * <p>Проверяет что столбец содержит формулу</p>
         * @return Класс-линкер для указания связи со следующим условием или перехода далее по алгоритму
         */
        @Contract(" -> new")
        public @NonNull LINKER isFormula() {
            return test(IS_FORMULA);
        }
        /**
         * Добавляет в инструкцию для парсера проверку ранее указанного столбца на несоответствие типу {@link CellType#FORMULA}
         * <p>Проверяет что столбец не содержит формулу</p>
         * @return Класс-линкер для указания связи со следующим условием или перехода далее по алгоритму
         */
        @Contract(" -> new")
        public @NonNull LINKER isNotFormula() {
            return test(IS_FORMULA.negate());
        }
        /**
         * Добавляет в инструкцию для парсера проверку ранее указанного столбца на тип {@link CellType#ERROR}
         * <p>Проверяет что столбец содержит формулу с ошибкой</p>
         * @return Класс-линкер для указания связи со следующим условием или перехода далее по алгоритму
         */
        @Contract(" -> new")
        public @NonNull LINKER isError() {
            return test(IS_ERROR);
        }
        /**
         * Добавляет в инструкцию для парсера проверку ранее указанного столбца на несоответствие типу {@link CellType#ERROR}
         * <p>Проверяет что столбец не содержит формулу с ошибкой</p>
         * @return Класс-линкер для указания связи со следующим условием или перехода далее по алгоритму
         */
        @Contract(" -> new")
        public @NonNull LINKER isNotError() {
            return test(IS_ERROR.negate());
        }
        /**
         * Добавляет в инструкцию для парсера проверку на {@code null}
         * <p>Проверяет что поле столбца является {@code null}</p>
         * @return Класс-линкер для указания связи со следующим условием или перехода далее по алгоритму
         */
        @Contract(" -> new")
        public @NonNull LINKER isNull() {
            return test(IS_NULL);
        }
        /**
         * Добавляет в инструкцию для парсера проверку на {@code null}
         * <p>Проверяет что поле столбца <b>НЕ</b> является {@code null}</p>
         * @return Класс-линкер для указания связи со следующим условием или перехода далее по алгоритму
         */
        @Contract(" -> new")
        public @NonNull LINKER isNotNull() {
            return test(IS_NULL.negate());
        }

        /**
         * Добавляет в инструкцию для парсера проверку содержимого ячейки на соответствие указанному значению
         * <p><b>АВТОМАТИЧЕСКИ</b> проверяет что ячейка имеет тип {@link CellType#STRING}</p>
         * @param other Строка, на соответствие которой надо проверить содержимое ячейки
         * @return Класс-линкер для указания связи со следующим условием или перехода далее по алгоритму
         */
        @Contract("_ -> new")
        public @NonNull LINKER stringValueEquals(@NonNull String other) {
            return test(IS_STRING.and(cell -> cell.getStringCellValue().equals(other)));
        }
        /**
         * Добавляет в инструкцию для парсера проверку содержимого ячейки на соответствие указанному значению без учета регистра
         * <p><b>АВТОМАТИЧЕСКИ</b> проверяет что ячейка имеет тип {@link CellType#STRING}</p>
         * @param other Строка, на соответствие которой надо проверить содержимое ячейки без учета регистра
         * @return Класс-линкер для указания связи со следующим условием или перехода далее по алгоритму
         */
        @Contract("_ -> new")
        public @NonNull LINKER stringValueEqualsIgnoreCase(@NonNull String other) {
            return test(IS_STRING.and(cell -> cell.getStringCellValue().equalsIgnoreCase(other)));
        }
        /**
         * Добавляет в инструкцию для парсера проверку содержимого ячейки на содержание указанного значения
         * <p><b>АВТОМАТИЧЕСКИ</b> проверяет что ячейка имеет тип {@link CellType#STRING}</p>
         * @param other Строка, на содержание которой надо проверить содержимое ячейки
         * @return Класс-линкер для указания связи со следующим условием или перехода далее по алгоритму
         */
        @Contract("_ -> new")
        public @NonNull LINKER stringValueContains(@NonNull String other) {
            return test(IS_STRING.and(cell -> cell.getStringCellValue().contains(other)));
        }
        /**
         * Добавляет в инструкцию для парсера проверку содержимого ячейки на содержание указанного значения без учета регистра
         * <p><b>АВТОМАТИЧЕСКИ</b> проверяет что ячейка имеет тип {@link CellType#STRING}</p>
         * @param other Строка, на содержание которой надо проверить содержимое ячейки без учета регистра
         * @return Класс-линкер для указания связи со следующим условием или перехода далее по алгоритму
         */
        @Contract("_ -> new")
        public @NonNull LINKER stringValueContainsIgnoreCase(@NonNull String other) {
            return test(IS_STRING.and(cell -> cell.getStringCellValue().toLowerCase().contains(other.toLowerCase())));
        }
    }

    /**
     * Класс описывает связку нескольких условий и предоставляет методы перехода к другому столбцу таблицы.
     * @param <CONDITION> Класс основного условия, которому нужны связки
     * @param <LINKER> Класс, в котором следует указать наследующий класс для выставления циклической связи
     */
    private static abstract class ConditionLinker<
            CONDITION extends Condition<LINKER, CONDITION>,
            LINKER extends ConditionLinker<CONDITION, LINKER>> {
        private final int cellNum;
        protected final Predicate<XSSFRow> initial;

        private ConditionLinker(int cellNum, @NonNull Predicate<XSSFRow> initial) {
            this.cellNum = cellNum;
            this.initial = initial;
        }

        /**
         * Возвращается к связанному классу {@link Condition}, связывая его переданным способом
         * @param cellNum Номер столбца, с которым производят работу
         * @param transformer Функция, применяемая к предыдущему условию, сочетающая его и условие, которое будет задано следующим вызовом
         * @return {@link Condition}, для дальнейшего набора условия
         */
        protected abstract CONDITION goBack(int cellNum, @NonNull Function<Predicate<XSSFRow>, @NonNull Predicate<XSSFRow>> transformer);

        /**
         * Аналог оператора {@code &&} для предыдущего и следующего условия.
         * <p>Имеет те же свойства, что и оператор {@code &&}:
         * если предыдущее условие не выполняется, проверка следующего не происходит, возвращается false</p>
         * @return Объект для задания следующего условия
         */
        public @NonNull CONDITION and() {
            return goBack(cellNum, initial::and);
        }
        /**
         * Аналог оператора {@code &&} для предыдущего и следующего условия. Следующее условие применяется к столбцу, на который указывает номер из аргумента
         * <p>Имеет те же свойства, что и оператор {@code &&}:
         * если предыдущее условие не выполняется, проверка следующего не происходит, возвращается false</p>
         * @param otherCellNum Номер столбца, на проверку которого надо перейти
         * @return Объект для задания следующего условия
         */
        public @NonNull CONDITION andOtherCell(int otherCellNum) {
            return goBack(otherCellNum, initial::and);
        }
        /**
         * Аналог оператора {@code ||} для предыдущего и следующего условия.
         * <p>Имеет те же свойства, что и оператор {@code ||}:
         * если предыдущее условие выполняется, проверка следующего не происходит, возвращается true</p>
         * @return Объект для задания следующего условия
         */
        public @NonNull CONDITION or() {
            return goBack(cellNum, initial::or);
        }
        /**
         * Аналог оператора {@code ||} для предыдущего и следующего условия. Следующее условие применяется к столбцу, на который указывает номер из аргумента
         * <p>Имеет те же свойства, что и оператор {@code ||}:
         * если предыдущее условие не выполняется, проверка следующего не происходит, возвращается true</p>
         * @param otherCellNum Номер столбца, на проверку которого надо перейти
         * @return Объект для задания следующего условия
         */
        public @NonNull CONDITION orOtherCell(int otherCellNum) {
            return goBack(otherCellNum, initial::or);
        }
    }
}