package com.kanayaya.XLSParse.InnerClassImplementation;

@FunctionalInterface
public interface UncheckedSupplier<T> {
    T get() throws Exception;
    default T getUnchecked() {
        try {
            return get();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}