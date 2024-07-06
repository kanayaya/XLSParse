package com.kanayaya.XLSParse.InnerClassImplementation;

@FunctionalInterface
public interface UncheckedConsumer<T> {
    void accept(T t) throws Exception;
    default void acceptUnchecked(T t) {
        try {
            accept(t);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}