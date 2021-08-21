# wow-excel
让读写 excel 更简单。

```xml
<dependency>
    <groupId>com.github.nekolr</groupId>
    <artifactId>wow-excel-poi</artifactId>
    <version>1.0.5</version>
</dependency>
```

# 特性
- 支持单级表头和多级表头
- 支持流式读取和流式写入（可以避免 OOM）
- 支持数据转换
- 支持自定义事件监听器（提供单元格级别、行级别、sheet 级别的监听器）
- 支持自定义样式监听器

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
@Excel(value = "计划书", type = WorkbookType.XLS)
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

以上演示的只是部分读的功能，其余功能可以查看文档。下面演示写的功能，写文档可能会自定义一些样式，因此我们可以重新定义一下实体，添加一些样式相关的配置：

```java
@Excel(value = "计划书", type = WorkbookType.XLSX)
@ToString
public class Project {

    @ExcelField(value = {"项目", "项目"}, width = 4096)
    private String project;

    @ExcelField({"负责人", "负责人"})
    private String manager;

    @ExcelField(value = {"立项时间", "立项时间"}, width = 4096, format = "m/d/yy")
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

其中设置了一些表头所在列的宽度，同时还设置了日期类型单元格的格式。接下来是写 excel 的代码：

```java
public class Test {
    public static void main(String[] args) throws FileNotFoundException {
        // 读数据
        List<Project> list = WoWExcel.createReaderBuilder("计划书.xlsx", Project.class)
                .rowNum(2)
                .build()
                .readAndGet();
        
        // 写数据
        FileOutputStream out = new FileOutputStream(new File("新计划书.xlsx"));
        WoWExcel.createWriterBuilder(out, Project.class)
                .enableMultiHead() // 开启写多级表头，如果不开启则不会合并多级表头
                .enableStreamingWriter() // 开启流式写
                .defaultStyle() // 使用默认提供的一套样式
                .build()
                .writeBigTitle(new BigTitle(2, 0, 6, "计划书")) // 写大标题
                .writeHead() // 写表头
                .write(list) // 写数据
                .flush(); // 最后一定要调用刷新的方法
    }
}
```

![写文档](https://github.com/nekolr/wow-excel/blob/master/media/write_multi_head_example.png)
