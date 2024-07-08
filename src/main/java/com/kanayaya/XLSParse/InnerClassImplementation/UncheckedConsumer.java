package com.kanayaya.XLSParse.InnerClassImplementation;

import lombok.NonNull;

import java.util.function.Consumer;
/**
 * То же, что и {@link Consumer}, но без обработки исключений.
 * @param <T> тип принимаемого значения.
 */
@FunctionalInterface
public interface UncheckedConsumer<T> extends Consumer<T> {
    /**
     * @param t Принимаемое значение.
     * @throws Exception в случае ошибки.
     */
    void acceptUnchecked(T t) throws Exception;

    /**
     * Принимает значение без обработки исключений. В случае ошибки перебрасывает, заворачивая в {@link RuntimeException}.
     * @param t Принимаемое значение.
     */
    default void accept(T t) {
        try {
            acceptUnchecked(t);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @NonNull
    @Override
    default UncheckedConsumer<T> andThen(@NonNull Consumer<? super T> after) {
        return (T t) -> { accept(t); after.accept(t); };
    }
}