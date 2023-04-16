package io.github.mbenincasa.javaexcelutils.model.zip;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;

/**
 * This is a model which contains the data of the file to be zipped which is used in {@code ZipUtility}
 * @author Mirko Benincasa
 * @since 0.4.2
 */
@AllArgsConstructor
@Getter
@Setter
public class ZipModel {

    /**
     * The Byte array of the file to zip
     */
    private byte[] bytes;

    /**
     * The name of the file to zip
     */
    private String filename;
}
