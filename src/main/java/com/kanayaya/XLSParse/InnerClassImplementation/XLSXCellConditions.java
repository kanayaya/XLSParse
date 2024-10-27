package com.kanayaya.XLSParse.InnerClassImplementation;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.util.Map;
import java.util.function.Function;

public class XLSXCellConditions {
    private static final Map<CellType, Function<CellValue, String>> stringConverters = Map.of(
            CellType.STRING, CellValue::getStringValue,
            CellType.BLANK, cellValue -> "",
            CellType.NUMERIC, cellValue -> Double.toString(cellValue.getNumberValue()),
            CellType.BOOLEAN, cellValue -> Boolean.toString(cellValue.getBooleanValue()),
            CellType.ERROR, CellValue::formatAsString
    );
    private final XSSFCell cell;
    private final FormulaEvaluator evaluator;

    public XLSXCellConditions(XSSFCell cell, FormulaEvaluator evaluator) {
        this.cell = cell;
        this.evaluator = evaluator;
    }
    public boolean isNumeric() {
        return cell != null && cell.getCellType().equals(CellType.NUMERIC);
    }
    public boolean isString() {
        return cell != null && cell.getCellType().equals(CellType.STRING);
    }
    public boolean isBlank() {
        return cell != null && cell.getCellType().equals(CellType.BLANK);
    }
    public boolean isBoolean() {
        return cell != null && cell.getCellType().equals(CellType.BOOLEAN);
    }
    public boolean isError() {
        return cell != null && cell.getCellType().equals(CellType.ERROR);
    }
    public boolean isFormula() {
        return cell != null && cell.getCellType().equals(CellType.FORMULA);
    }
    public String rawValue() {
        return cell.getRawValue();
    }
    public String stringValue() {
        return cell.getStringCellValue();
    }
    public String stringEvaluatedValue() {
        CellValue evaluate = evaluator.evaluate(cell);
        return stringConverters.getOrDefault(evaluate.getCellType(), CellValue::formatAsString).apply(evaluate);
    }
}
