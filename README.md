# ExcelDna.CellAddress
Excel Range Warpper using C#

[![Build status](https://ci.appveyor.com/api/projects/status/ykuyq8exknf4kojn/branch/master?svg=true)](https://ci.appveyor.com/project/zwq000/exceldna-celladdress/branch/master)

Excel Range Com 对象和 ExcelDna.ExcelReference 都需要Excel 作为运行环境才能使用，
如果需要在其他环境使用单元格地址，那么你需要一个 可以离线使用的版本。
完成基本的单元格地址空间计算，同时还能在连线时直接访问 Excel 对象。


## 使用方式
```
PS> Install-Package  ExcelDna.ExcelAddress
```

## 对象构造

- 字符串转换为单元格地址
```C#
    //单个单元格
    var cell = CellAddress.Parse("Sheet 1!A1")
    //单元格范围
    var range = CellAddress.Parse("Sheet 1!A1:F4")

    //R1C1 格式
    var rangeR1C1 = CellAddress.Parse("Sheet 1!R1C1:R4C4")

```

- Range 对象构造

```C#
    Range range;
    var cell = new CellAddress(range)
```

- ExcelDna Reference 对象构造

```C#
    ExcelReference range;
    var cell = new CellAddress(range)
```

- 隐式转换
```C#
    var cell = (CellAddress)"A1"
```


## 扩展方法

### GetValue 
  > 获取地址引用单元格 的数据
  > 支持泛型方法

### SetValue
  > 向 地址引用单元格 写入数据

###  HasFormula
  > 地址所在单元格是否包含公式

###  SetFormula
  > 设置公式

###  单元格区域 内部单元格遍历方法
  - GetCells
    > 根据索引返回区域内部的第n个单元格

  - Offset 
    > 计算给定单元格地址的偏移。返回的区域和原始区域大小相同


### Min
  > 两个/多个单元格地址比较大小, 右下方较大

### Max 
  > 两个/多个单元格地址中最大地址

### GetRange
  > 计算多个单元格地址所占用的区域