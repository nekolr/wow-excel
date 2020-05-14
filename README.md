# wow-excel
让读写 excel 更简单。

# 特性
- 支持单级表头和多级表头
- 支持流式读取和流式写入（可以避免 OOM）
- 支持数据转换
- 支持自定义事件监听器

# 快速上手
首先需要新建一个实体类，实体类的作用一是方便表格数据的存储；二是用来描述表格的一些信息。比如有这样一个表格：

![单级表头](https://github.com/nekolr/wow-excel/blob/master/media/read_single_head_example.png)

那么对应的实体就是这样的：

```java
@Excel("计划书")
@ToString
public class Project {
    @ExcelField("项目")
    private String project;
    @ExcelField("负责人")
    private String manager;
    @ExcelField("立项时间")
    private Date date;
    @ExcelField("专项资金")
    private Long money;
}
```

接下来就是读 excel 的代码：

```java
public class Test {
    public static void main(String[] args) {
        List<Project> list = WoWExcel.createReaderBuilder("计划书.xls", Project.class)
                .rowNum(2)
                .colNum(1)
                .build()
                .readAndGet();
        list.forEach(System.out::println);
    }
}
```

有时候表头可能有多级，比如这样一个 excel 表格：

![多级表头](https://github.com/nekolr/wow-excel/blob/master/media/read_multi_head_example.png)

那么对应的实体应该是这样的：

```java
@Excel(value = "计划书", type = ExcelType.XLS)
@ToString
public class Project {

    @ExcelField({"项目", "项目"})
    private String project;

    @ExcelField({"负责人", "负责人"})
    private String manager;

    @ExcelField({"立项时间", "立项时间"})
    private Date date;

    @ExcelField({"年度预算", "一季度"})
    private Double q1;

    @ExcelField({"年度预算", "二季度"})
    private Double q2;

    @ExcelField({"年度预算", "三季度"})
    private Double q3;

    @ExcelField({"年度预算", "四季度"})
    private Double q4;
}
```

通过这种方式构建的多级表头，能够在定义之初就将整个表头部分的结构描述出来，这样在写 excel 的时候再进行行合并或者列合并就比较简单了。下面是读取 excel 的代码：

```java
public class Test {
    public static void main(String[] args) {
        WoWExcel.createReaderBuilder("计划书.xls", Project.class)
                .rowNum(2)
                .subscribe(list -> list.forEach(System.out::println))
                .build()
                .read();
    }
}
```