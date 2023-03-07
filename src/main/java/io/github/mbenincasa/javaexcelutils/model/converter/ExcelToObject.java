package io.github.mbenincasa.javaexcelutils.model.converter;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;

@AllArgsConstructor
@Getter
@Setter
public class ExcelToObject<T> {

    private String sheetName;
    private Class<T> clazz;
}
