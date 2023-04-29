package io.github.mbenincasa.javaexcelutils.tools;

import io.github.mbenincasa.javaexcelutils.model.zip.ZipModel;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * {@code ZipUtility} is a static class that provides utility methods for zipped files
 * @author Mirko Benincasa
 * @since 0.4.2
 */
public class ZipUtility {

    /**
     * @param zipModels A list of ZipModel objects representing the files to zip
     * @return The zipped file in the form of a ByteArrayOutputStream
     * @throws IOException If an I/O error has occurred
     */
    public static ByteArrayOutputStream zipFiles(List<ZipModel> zipModels) throws IOException {
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        ZipOutputStream zipOutputStream = new ZipOutputStream(byteArrayOutputStream);
        for (ZipModel zipModel : zipModels) {
            ZipEntry zipEntry = new ZipEntry(zipModel.getFilename());
            zipEntry.setSize(zipModel.getBytes().length);
            zipOutputStream.putNextEntry(zipEntry);
            zipOutputStream.write(zipModel.getBytes());
        }

        zipOutputStream.closeEntry();
        zipOutputStream.close();
        return byteArrayOutputStream;
    }
}
