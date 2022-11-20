package enums;

import lombok.AllArgsConstructor;
import lombok.Getter;

@Getter
@AllArgsConstructor
public enum Extension {

    XLS("xls"),
    XLSX("xlsx");

    private final String ext;
}
