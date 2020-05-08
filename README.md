# wow-excel
让读写 excel 更简单。

# 特性
- 支持单级表头和多级表头
- 支持数据转换
- 支持自定义空值处理方法
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
        List<Project> list = WoWExcel.createReader("计划书.xls", Project.class)
                .rowIndex(2)
                .colIndex(1)
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
public class Demo {

    @ExcelField(value = "项目")
    private String project;

    @ExcelField(value = "负责人")
    private String manager;

    @ExcelField(value = "立项时间")
    private Date date;

    @ExcelField(value = "一季度")
    private Double q1;

    @ExcelField(value = "二季度")
    private Double q2;

    @ExcelField(value = "三季度")
    private Double q3;

    @ExcelField(value = "四季度")
    private Double q4;
}
```

接下来是读取 excel 的代码：

```java
public class Test {
    public static void main(String[] args) {
        List<Project> list = WoWExcel.createReader("计划书.xls", Project.class)
                .rowIndex(2)
                .readAndGet();
        list.forEach(System.out::println);
    }
}
```