package com.kanayaya.XLSParse.InnerClassImplementation;

public interface UncheckedBiConsumer<T, U> {
    void accept(T u, U t) throws Exception;
    default void acceptUnchecked(T t, U u) {
        try {
            accept(t, u);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}