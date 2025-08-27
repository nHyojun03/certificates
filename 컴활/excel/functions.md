# Excel 함수

## <span style="color: #083891">날짜, 시간 관련</span>

※ `date`와 `datetime`, `time` 은 `number`와 같은 타입

### 현재
- `=NOW()` => `date`
   : 현재 날짜, 시간
- `=TODAY()` => `date`
   : 현재 날짜

### 시간 입력
- `=DATE(year: number, month: number, day: number)`=> `date`
   : 날짜 입력
- `=DATEVALUE(date_text: string)` => `datetime`
   : 서식으로 날짜 입력
- `=TIME(hour: number, minute: number, second: number)` => `time`
   : 시간 입력

### 날짜 데이터 접근
- `=YEAR(serial: datetime)` => `number`
- `=MONTH(serial: datetime)` => `number`
- `=DAY(serial: datetime)` => `number`
- `=HOUR(serial: datetime)` => `number`
- `=MINUTE(serial: datetime)` => `number`
- `=SECOND(serial: datetime)` => `number`

#### [=WEEKDAY(serial, [return_type: 1])](https://support.microsoft.com/ko-kr/office/weekday-%ED%95%A8%EC%88%98-60e44483-2ed1-439f-8bd0-e404c190949a)

: 무슨 요일?

- serial: date
- return_type

| return_type | 시작 | 끝 |
| :---: | :---: | :---: |
| 1 or 11 | 1(日) | 7(土) |
| 2 | 1(月) | 7(日) |
| 3 | 0(月) | 6(日) |

그 외 12, 13, 14, 15, 16, 17 모두 1~7을 반환


#### [=WEEKNUM(serial, [return_type: 1])](https://support.microsoft.com/ko-kr/office/weeknum-%ED%95%A8%EC%88%98-e5c43a03-b4ab-426c-b411-b18c13c75340)

: 몇 번째 주?

- serial: date
- return_type

| return_type | 주가 시작되는 요일 |
| :---: | :---: | 
| 1 | 日 |
| 2 | 月 |
| 11 | 月 |
| 12 | 火 |
| ... | |
| 17 | 日 |
| 21 | 月 |

### 날짜 계산

- `=DAYS(end: date, start: date)` => `number`
   : `end` - `start`
- [`=NETWORKDAYS(start: date, end: date, [holidays: array<date>])`](https://support.microsoft.com/ko-kr/office/networkdays-%ED%95%A8%EC%88%98-48e717bf-a7a3-495f-969e-5005e3eb18e7) => `number`
   : `end` - `start` - n(`holidays`)
- [`=EDATE(start: date, months: number)`](https://support.microsoft.com/ko-kr/office/edate-%ED%95%A8%EC%88%98-3c920eb2-6e66-44e7-a1f5-753ae47ee4f5) => `number`
   : `start` + `months`개월
- [`=EOMONTH(start: date, months: number)`](https://support.microsoft.com/ko-kr/office/eomonth-%ED%95%A8%EC%88%98-7314ffa1-2bc9-4005-9d66-f49db127d628) => `number`
   : `start` + `months`개월, 해당 달의 마지막 날짜 (30일, 31일 등등)
- [`=WORKDAY(start: date, days: number, [holidays: array<date>])`](https://support.microsoft.com/ko-kr/office/workday-%ED%95%A8%EC%88%98-f764a5b7-05fc-4494-9486-60d494efbf33) => `number`
   : `start` + `date`. 근데 중간에 `holidays`가 있으면 그걸 빼고 계산

---

## <span style="color: #083891">논리</span>

- `=TRUE()`
- `=FALSE()`


- `=AND(...![logic])`
- `=OR(...![logic])`
- `=NOT(logic)`


- `=IF(logic, value_if_true, [value_if_false])`
  - `value_if_false`가 정의되지 않은 경우 `FALSE`를 반환
- `=IFS(...![logic, value_if_true])`
  - 어떤 `logic`도 `TRUE`가 아닌 경우 `#N/A`를 반환
- `=IFERROR(value, value_if_error)`
- `=SWITCH(exp, ...![value, result], [default])`
  - `exp` = `value`가 성립되는 값이 없고, `default`도 없으면 `#N/A` 반환

---

## <span style="color: #083891">문자열</span>

- `=LEN(text: string)` => `number`
   : 문자수

### 문자열 비교

- `=EXACT(text1: string, text2: string)` => `boolean`
   : 문자열이 같으면 TRUE, 다르면 FALSE

### 영문자 대소문자 관련

- `=LOWER(text: string)` => `string`
   : 모든 영문자를 소문자로
- `=PROPER(text: string)` => `string`
   : 첫번째와 영문자가 아닌 문자 다음에 오는 첫번째 영문자를 대문자로, 
     &nbsp; 그 외는 소문자로
- `=UPPER(text: string)` => `string`
   : 모든 영문자를 대문자로

### 문자열 연산

- `=CONCAT(...![text: string])` => `string`
- `=REPT(text: string, times: number)` => `string`
- `=TRIM(text: string)` => `string`
- 문자열 자르기
  - `=LEFT(text: string, [nchars: number(1)])` => `string`
  - `=MID(text: string, start: number, nchars: number)` => `string`
  - `=RIGHT(text: string, [nchars: number(1)])` => `string`
- 부분 문자열 수정
  - [`=REPLACE(old: string, start: num, nchars: num, new: string)`](https://support.microsoft.com/ko-kr/office/replace-%ED%95%A8%EC%88%98-8d799074-2425-4a8a-84bc-82472868878a)
     : `start` 부터 `nchars`문자의 부분 문자열을 `new`로 바꿈
  - [`=SUBSTITUTE(text: str, old: str, new: str, [instance: num])`](https://support.microsoft.com/ko-kr/office/substitute-%ED%95%A8%EC%88%98-6434944e-a904-4336-a9b0-1e58df3bc332)
     : `text` 속 `instance`번째 `old`를 `new`로 대체.
     &nbsp;`instance`가 생략되어 있으면 모든 `old`를 `new`로 대체

### 문자열 검색

: `within_text`의 `start`번째 문자에서 부터 찾은 `find_text`의 시작 위치

- `=FIND(find_text, within_text, [start: num(1)])` => `number`
- `=SEARCH(find_text, within_text, [start: num(1)])` => `number`

| | FIND | SEARCH |
| :---: | :---: | :---: |
| 대소문자 구분 | O | X |
| 와일드 문자 지원 | X | O | 

### 문자열 ⇔ 숫자

- `=TEXT(value: number, format_text: string)` => `string`
   : `number` to `string`
   - 대소문자 
- `=VALUE(text: string)` => `number`
   : `string` to `number`

---

## <span style="color: #083891">수학</span>

- `=PI()`: $\pi$
- `=SIGN(n)`: $\operatorname{sgn}{x}$
- `=ABS(n)`: $\left | n \right |, n \in \mathbb{R}$

### 올림, 내림, 반올림

- `=INT(n)` : $\left\lfloor n\right\rfloor$
- `=TRUNC(n)`: 소수점 아래 그냥 버림. 
  - `INT(n)`와 음수에서 차이남
- `=ROUND(n, d)` : `n`을 소수점 아래 `d`번째 자리로 **반올림**
- `=ROUNDUP(n, d)`: 올림
- `=ROUNDDOWN(n, d)`: 내림
- ※ [`=FIXED(n, [d: 2], [no_commas: FALSE])`](https://support.microsoft.com/ko-kr/office/fixed-%ED%95%A8%EC%88%98-ffd5723c-324c-45e9-8b96-e41be2a8274a) => `string`
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : 숫자 서식 지정 

### 난수

- `=RAND()`: 0~1 사이의 난수
- `=RANDBETWEEN(a, b)`: a~b 사이의 정수인 난수

### 연산

- `=SUM(...[n: number | array<number>])`: $\sum{n_i}$
- [SUMIF](https://support.microsoft.com/ko-kr/office/sumif-%ED%95%A8%EC%88%98-169b8c99-c05c-4483-a712-1697a653039b)
- [SUMIFS](https://support.microsoft.com/ko-kr/office/sumifs-%ED%95%A8%EC%88%98-c9e748f5-7ea7-455d-9406-611cebce642b)
- [SUMPRODUCT](https://support.microsoft.com/ko-kr/office/sumproduct-%ED%95%A8%EC%88%98-16753e75-9f68-4874-94ac-4d2145a2fd2e)
- `=FACT(n)`: $n!$
- `=PRODUCT(...[n: number | array<number>])`: $\prod{n_i}$
- `=QUOTIENT(a, b)` : $\left ( a \div b \right )$의 몫
- `=MOD(a, b)`: $\left ( a \div b \right )$의 나머지
- `=POWER(a, b)` = `a^b`
- `=EXP(n)`: $e^n$
- `=SQRT(n)`: $\sqrt{n}$

### 행렬 연산

- `=MMULT(m1, m2)`: 행렬곱(Matrix Multiplication)
- `=MDETERM(m)`: 행렬식(Matrix Determinant)
- `=MINVRSE(m)`: 역행렬(Inverse matrix)

---

## <span style="color: #083891">데이터베이스</span>

※ **파라미터**

- `database`: 레코드와 필드
- `field`: 연산에 사용되는 필드
- `criteria`: 조건 테이블

※ **함수**
- [`=DAVERAGE(db, field, criteria)`](https://support.microsoft.com/ko-kr/office/daverage-%ED%95%A8%EC%88%98-a6a2d5ac-4b4b-48cd-a1d8-7b37834e5aee): 평균
- [`DCOUNT(db, field, criteria)`](https://support.microsoft.com/ko-kr/office/dcount-%ED%95%A8%EC%88%98-c1fc7b93-fb0d-4d8d-97db-8d5f076eaeb1): 숫자값이 있는 셀의 개수
- [`DCOUNTA(db, field, criteria)`](https://support.microsoft.com/ko-kr/office/dcounta-%ED%95%A8%EC%88%98-00232a6d-5a66-4a01-a25b-c1653fda1244): 비어 있지 않은 셀의 개수
- [`DGET(db, field, criteria)`](https://support.microsoft.com/ko-kr/office/dget-%ED%95%A8%EC%88%98-455568bf-4eef-45f7-90f0-ec250d00892e): 한 셀의 값
  - `criteria`의 조건에 맞는 값이 없으면 `#VALUE` 오류 값 반환
  - 여러 값이면 `#NUM!` 오류 값 반환
- [`DMAX(db, field, criteria)`](https://support.microsoft.com/ko-kr/office/dmax-%ED%95%A8%EC%88%98-f4e8209d-8958-4c3d-a1ee-6351665d41c2): 최댓값
- [`DMIN(db, field, criteria)`](https://support.microsoft.com/ko-kr/office/dmin-%ED%95%A8%EC%88%98-4ae6f1d9-1f26-40f1-a783-6dc3680192a3): 최솟값
- [`DPRODUCT(db, field, criteria)`](https://support.microsoft.com/ko-kr/office/dproduct-%ED%95%A8%EC%88%98-4f96b13e-d49c-47a7-b769-22f6d017cb31): 곱
- [`DSTDEV(db, field, criteria)`](https://support.microsoft.com/ko-kr/office/dstdev-%ED%95%A8%EC%88%98-026b8c73-616d-4b5e-b072-241871c4ab96): 표준편차
- [`DSUM(db, field, criteria)`](https://support.microsoft.com/ko-kr/office/dsum-%ED%95%A8%EC%88%98-53181285-0c4b-4f5a-aaa3-529a322be41b): 합계
- [`DVAR(db, field, criteria)`](https://support.microsoft.com/ko-kr/office/dvar-%ED%95%A8%EC%88%98-d6747ca9-99c7-48bb-996e-9d7af00f3ed1): 분산

---

## <span style="color: #083891">통계</span>

- `=AVERAGE(...[number])`
- `=AVERAGEA(...[value])`
- `=AVERAGE(...[number])`


---

## <span style="color: #083891">정보</span>

---

## <span style="color: #083891">재무</span>

---

## <span style="color: #083891">찾기와 참조</span>
