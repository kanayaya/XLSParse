package com.kanayaya.XLSParse.InnerClassImplementation;

import lombok.NonNull;

import java.util.Objects;
import java.util.function.BiConsumer;
/**
 * То же, что и {@link BiConsumer}, но без обработки исключений.
 * @param <T> Первый тип параметра.
 * @param <U> Второй тип параметра.
 */
public interface UncheckedBiConsumer<T, U> extends BiConsumer<T, U> {
    /**
     * @param u Первый параметр.
     * @param t Второй параметр.
     * @throws Exception в случае ошибки.
     */
    void acceptUnchecked(T u, U t) throws Exception;

    /**
     * Принимает значения без обработки исключений. В случае ошибки перебрасывает, заворачивая в {@link RuntimeException}.
     * @param t Первый параметр.
     * @param u Второй параметр.
     */
    default void accept(T t, U u) {
        try {
            acceptUnchecked(t, u);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
    @NonNull
    @Override
    default UncheckedBiConsumer<T, U> andThen(@NonNull BiConsumer<? super T, ? super U> after) {
        Objects.requireNonNull(after);

        return (l, r) -> {
            accept(l, r);
            after.accept(l, r);
        };
    }
}