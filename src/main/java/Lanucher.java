import com.suk.poihelper.excelhelper.util.XlsUtil;
import entity.Worker;

import java.util.ArrayList;

public class Lanucher {

    public static void main(String[] args){
        ArrayList<Worker> workers = new ArrayList<>();
        Worker worker = new Worker();
        worker.setName("工人1");
        worker.setAge(25);
        worker.setGender("男");
        workers.add(worker);
        XlsUtil.export("C:/Users/HP/Desktop/test.xls",workers);
    }
}
