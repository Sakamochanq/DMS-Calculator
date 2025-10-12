<div align="center">
    <a href="#">
        <img src="../assets/DMS-Calculator-Logo.png" width="500px">
    </a>
    <br>
    <hr>
</div>

<br>

[日本語](../README.md)　｜　[English](./README-en.md)　｜　한국어

<br>

일반적으로 Office Excel은 셀에서 DMS 계산을 수행할 수 없습니다. 여러 셀과 복잡한 변환 수식을 사용하여 계산을 수행하는 것은 가능하지만 매우 어렵습니다. 따라서 셀에서 DMS 계산을 수행하면서도 함수 계산을 계속할 수 있는 VBA 스크립트를 개발했습니다. 간단한 사용법은 다음 데모를 확인해 주십시오. 🌵

<br>

<div align="center">
    <a href="#">
        <img src="../assets/DMS-Calculator-Demo.gif" width="450px">
    </a>
</div>

<br>
<br>

셀에 다음 접두사를 입력하면 다양한 기능을 자동 완성할 수 있습니다.
```
=sok_
```

<br>

### 왜 sok?

測量 = sokuryou = sok

<br>

### 기능
- [x] 추가 및 범위 지정.
- [ ] 음수 처리 간소화.
- [x] 방위각 논리.
- [x] 코사인, 사인 계산.
- [x] 나침반 추가.
- [x] 문자열을 십진수로 변환 함수.
- [ ] 리팩토링.

<br>

### 사용법

1. 리포지토리 복제

```
git clone https://github.com/Sakamochanq/DMS-Calculator.git
```

<br>

2. 엑셀에서 개발 탭을 통해 VBA 편집기를 실행합니다.

3. `./src/main.vba` 파일을 원하는 워크시트에 추가합니다.

4. 저장 후 해당 셀에서 **sok func**를 사용할 수 있습니다.

5. 매크로는 일반적인 xlsx 형식 대신 매크로 사용 가능 형식(*.xlsm)으로 저장하여 언제든지 실행할 수 있습니다.

<br>

> [!Note]  
> 이 스크립트는 Excel 및 VBA와 함께 작동하도록 설계되었으며, 다른 환경에서는 제대로 작동하지 않을 수 있습니다.

<br>

### 예시

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

### 저자

- 개발 [Sakamochanq](https://github.com/Sakamochanq)

- 기여 [Github Copilot](https://github.com/features/copilot)

- 번역 [DeepL](https://www.deepl.com/)

<br>

### 라이선스

Release under the [MIT](https://github.com/Sakamochanq/DMS-Calculator/blob/master/LICENSE) LICENSE
