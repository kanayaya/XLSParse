package com.kanayaya.XLSParse.InnerClassImplementation;

import lombok.NonNull;

import java.util.function.BiFunction;

/**
 * Транзитивная функция, передающая первый параметр по всей цепочке, <br>
 * второй параметр передаваемый в цепочку является результатом её выполнения
 * @param <T> Транзитивно передаваемый в цепочке тип
 * @param <C> Тип, меняющийся по цепочке
 * @param <R> Тип результата выполнения функции
 */
@FunctionalInterface
public interface TransitiveBiFunction<T, C, R> extends BiFunction<T, C, R> {
    /**
     * @param that Транзитивная функция, которая выполнится следующей
     * @param <R2> Тип результата, возвращаемого следующей функцией
     * @return Спаренные в цепочку функции
     */
    default <R2> TransitiveBiFunction<T, C, R2> andThen(@NonNull BiFunction<? super T, ? super R, R2> that) {
        return (transitive, chained) -> that.apply(transitive, apply(transitive, chained));
    }
}