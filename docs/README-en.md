<div align="center">
    <a href="#">
        <img src="../assets/DMS-Calculator-Logo.png" width="500px">
    </a>
    <br>
    <hr>
</div>

<br>

[日本語](../README.md)　｜　English　｜　[한국어](./README-kr.md)

<br>

Normally Office Excel does not allow DMS calculations to be performed on cells. It is possible to do the calculation by using multiple cells and complicated conversion formulas, but it is very difficult. Therefore, we have developed a VBA Script that can perform DMS calculations on cells and still perform function calculations. Please check the following Demo for a simple usage. 🌵

<br>

<div align="center">
    <a href="#">
        <img src="../assets/DMS-Calculator-Demo.gif" width="450px">
    </a>
</div>

<br>
<br>

Various functions can be auto-completed by entering the following prefixes on the cells.
```
=sok_
```

<br>

### Why sok ?

測量 = sokuryou = sok

<br>

### Features
- [x] Add and range specification.
- [ ] Simplified handling of negative numbers.
- [x] Azimuth Logic.
- [x] Calculation of cos sin.
- [x] Add Compass
- [x] String to Decimal func.
- [ ] Refactoring.

<br>

### Usage

1.　Repository Clone

```
git clone https://github.com/Sakamochanq/DMS-Calculator.git
```

<br>

2.　From Excel launch VBAEditor from the Development tab.

3.　Add `./src/main.vba` to any Worksheet.

4.　Once saved, you can use the **sok func** on the cell.

5.　Macros can be executed at any time by saving in macro-enabled format(*.xlsm) instead of the usual xlsx format.

<br>

> [!Note]  
> This script is designed to work with Excel and VBA, and it may not function correctly in other environments.

<br>

### Example

```python
A1 = 179°50′0″
B1 = 0°10′0″

=sok_add(A1, B1) #180°0′0″
```

<br>

```python
A1 = 180°50′0″
B1 = 0°50′0″

=sok_sub(A1, B1) #180°0′0″
```

<br>

```python
A1 = 179°30′0″
A2 = 0°10′0″
A3 = 0°20′0″

=sok_sumAll(A1:A3) #180°0′0″
```

<br>

```python
A1 = 180°0′0″

=sok_compass(A1) #SE
```

<br>
<hr>

### Author 

- Developing by [Sakamochanq](https://github.com/Sakamochanq)

- Contributing by [Github Copilot](https://github.com/features/copilot)

- Translation by [DeepL](https://www.deepl.com/)

<br>

### License

Release under the [MIT](https://github.com/Sakamochanq/DMS-Calculator/blob/master/LICENSE) LICENSE
