package com.kanayaya.XLSParse.InnerClassImplementation;

import java.util.function.Supplier;

/**
 * То же, что и {@link Supplier}, но без обработки исключений.
 * @param <T> тип возвращаемого значения.
 */
@FunctionalInterface
public interface UncheckedSupplier<T> extends Supplier<T> {
    /**
     * @return возвращаемое значение.
     * @throws Exception в случае ошибки.
     */
    T getUnchecked() throws Exception;

    /**
     * Возвращает значение без обработки исключений. В случае ошибки перебрасывает, заворачивая в {@link RuntimeException}.
     * @return возвращаемое значение.
     */
    default T get() {
        try {
            return getUnchecked();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}