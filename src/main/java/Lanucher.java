import com.suk.poihelper.excelhelper.util.XlsUtil;
import entity.Worker;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Lanucher {

    public static void main(String[] args){
        ArrayList<Worker> workers = new ArrayList<>();
        Worker worker = new Worker();
        worker.setName("工人1");
        worker.setAge("25");
        worker.setGender("男");
        workers.add(worker);
        XlsUtil.export("C:\\Users\\Administrator\\Desktop\\test.xls",workers);

        HashMap<List<Object>, HashMap<Map<Integer, Integer>, String>> improt = XlsUtil.improt("C:\\Users\\Administrator\\Desktop\\test.xls", null, Worker.class);
    }
}
