import com.wixdom.JsonMergeXlsx;
import org.junit.jupiter.api.Test;

import java.io.IOException;

/**
 * @author Lin Qinghua
 */
public class JsonMergeXlsxTest {

    private static String JSON_FILENAME = "test.json";

    private static String XLSX_FILENAME = "test.xlsx";

    @Test
    public void test() throws IOException {
        String jsonPath = getJsonPath();
        String xlsxPath = getXslxPath();
        JsonMergeXlsx merger = new JsonMergeXlsx();
        merger.merge(jsonPath,"/b",xlsxPath,"Sheet1");
    }

    private String getXslxPath() {
        return getResoucePath(XLSX_FILENAME);
    }

    private String getJsonPath() {
        return getResoucePath(JSON_FILENAME);
    }

    private String getResoucePath(String resource) {
        return getClass().getClassLoader().getResource(resource).getPath();
    }

}
