package io.github.mbenincasa.javaexcelutils.tools;

import io.github.mbenincasa.javaexcelutils.model.zip.ZipModel;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.*;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

class ZipUtilityTest {

    private final File excelFile = new File("./src/test/resources/employee_2.xlsx");
    private final File jsonFile = new File("./src/test/resources/office.json");

    @Test
    void zipFiles() throws IOException {
        byte[] bytes = Files.readAllBytes(excelFile.toPath());
        ZipModel zipModel = new ZipModel(bytes, excelFile.getName());
        byte[] bytes1 = Files.readAllBytes(jsonFile.toPath());
        ZipModel zipModel1 = new ZipModel(bytes1, jsonFile.getName());

        List<ZipModel> zipModels = new ArrayList<>();
        zipModels.add(zipModel);
        zipModels.add(zipModel1);
        ByteArrayOutputStream byteArrayOutputStream = ZipUtility.zipFiles(zipModels);
        FileOutputStream fileOutputStream = new FileOutputStream("./src/test/resources/file.zip");
        fileOutputStream.write(byteArrayOutputStream.toByteArray());

        fileOutputStream.close();

        File file = new File("./src/test/resources/file.zip");
        FileInputStream fs = new FileInputStream(file);
        ZipInputStream Zs = new ZipInputStream( new BufferedInputStream(fs));

        ZipEntry ze = Zs.getNextEntry();
        assert ze != null;
        Assertions.assertEquals("employee_2.xlsx", ze.getName());
        ze = Zs.getNextEntry();
        Assertions.assertEquals("office.json", ze.getName());

        fs.close();
        file.delete();
    }
}