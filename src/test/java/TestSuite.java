import model.ExcelWorkbookTest;
import org.junit.platform.suite.api.SelectClasses;
import org.junit.platform.suite.api.Suite;
import tools.ConverterTest;
import tools.ExcelUtilityTest;
import tools.SheetUtilityTest;
import tools.WorkbookUtilityTest;

@Suite
@SelectClasses({ConverterTest.class, ExcelUtilityTest.class, SheetUtilityTest.class, WorkbookUtilityTest.class, ExcelWorkbookTest.class})
public class TestSuite {
}
