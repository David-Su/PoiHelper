package entity;

import com.suk.poihelper.excelhelper.annotation.ExcelFileAttr;
import com.suk.poihelper.excelhelper.annotation.ExcelTypeAttr;

@ExcelTypeAttr(titleStr = "工人")
public class Worker {

    @ExcelFileAttr(nameStr = "姓名",column = 0)
    private String name;

    @ExcelFileAttr(nameStr = "性别",column = 1)
    private String gender;

    @ExcelFileAttr(nameStr = "年龄",column = 2)
    private int age;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getGender() {
        return gender;
    }

    public void setGender(String gender) {
        this.gender = gender;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }
}
