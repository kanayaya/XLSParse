package com.kanayaya.XLSParse.InnerClassImplementation;

import lombok.NonNull;

import java.util.function.BiFunction;


@FunctionalInterface
public interface TransitiveBiFunction<T, C, R> extends BiFunction<T, C, R> {
    default <R2> TransitiveBiFunction<T, C, R2> andThen(@NonNull TransitiveBiFunction<? super T, ? super R, R2> that) {
        return (transitive, chained) -> that.apply(transitive, apply(transitive, chained));
    }
}