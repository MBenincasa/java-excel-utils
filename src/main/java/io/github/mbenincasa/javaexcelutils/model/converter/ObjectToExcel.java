package io.github.mbenincasa.javaexcelutils.model.converter;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;

import java.util.stream.Stream;

@AllArgsConstructor
@Getter
@Setter
public class ObjectToExcel<T> {

    private String sheetName;
    private Class<T> clazz;
    private Stream<T> stream;
}
