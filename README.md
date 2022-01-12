# projectJH

* 파이썬으로 크롤링 하기
* 메인 라이브러리: PyQt5, Selenium
* URL: https://www.maersk.com/schedules
* 목적: 인턴 중인 박주환 도와주기
* 기능: 특정 날짜의 특정 선박의 스케쥴을 크롤링 해오기, 일정 간격으로 자동 업데이트 가능


## ver 1.0 

### 실행 화면

<img width="216" alt="jhresult" src="https://user-images.githubusercontent.com/78152114/141646688-a042a7bb-4721-4ae0-847a-ef6c634035cf.png">

<img width="216" alt="jhresult" src="https://user-images.githubusercontent.com/78152114/141647260-27852112-fd22-4485-b3c0-ac3b6a0e616a.png">

Last Updated -2021/11/13 SAT
1. 선박 스케쥴 결과 표시
2. 수동 업데이트
3. 업데이트 정보 표시
4. 자동 업데이트

----

## ver 2.0

### 실행화면
<img width="400" alt="스크린샷 2022-01-12 오후 3 36 25" src="https://user-images.githubusercontent.com/78152114/149076654-1a771abc-02b7-4bf7-805d-2532fe529ef3.png">

### 설명

엑셀시트 내의 두 개의 `테이블`을 찾아 `데이터프레임`으로 변환 후 `테이블 위젯`으로 변경 (Excel -> DataFrame -> TableWidget)

중간의 DataFrame은 `임시 저장소` 역할이 가능하다. 

[머스크 페이지](https://www.maersk.com/schedules/vesselSchedules) 에서 찾고자 하는 `vessel`를 검색해 스케쥴을 검색한다.

URL 검색을 통해 이루어지며 vessel code와 date가 필요하다.

vessel code는 `매핑 테이블`을 직접 만들어 관리하며 date는 원하는 날짜 기준으로 7일 간격으로 검색한다.

두 개의 테이블은 서로 다른 두 개의 항로임을 나타내며, tabel1은 `Bremerhaven`과 `Gdansk`항을 , table2는 `Bremerhaven`항만 필요하다.

이미 정해진 항로를 바탕으로 찾고자 하는 항구의 이전의 항구를 검색하여 찾는다.

table1은 Bremerhaven -> Gdansk -> Bremerhaven 순이며, table2는 Bremerhaven -> Bremerhaven이다.

검색하는 시점의 상황 또는 선박 사정에 따라 일부 항구를 건너뛰는 경우, 아직 스케줄에 업데이트 되지 않은 경우 등이 존재하기 때문에 상황에 맞는 알고리즘이 필요하다.

알고리즘은 프로그램에서 테이블의 특정 셀을 클릭하면 해당 행에서 *Vessel Name*에 해당하는 열을 찾아 찾고자하는 선박 이름을 추출하고 이를 매핑 테이블에서 검색해 선박 코드를 찾는다.

찾은 선박 코드와 날짜를 url에 맞게 입력해 검색을 요청하고 원하는 데이터를 찾는다.

## Tips

### 엑셀에서 숨겨지지 않은 행들만 찾기

하나의 엑셀시트 내에서 서로 나눠진 두 개 테이블을 구하고 싶다.

각 테이블의 위의 열에는 숨김 처리된 행들이 있다.

이러한 행들을 제외한 행들만 구해 테이블의 원하는 값들만 구하고 싶다.

`ws.row_dimensions[i]`: 워크시트에서 원하는 i번째 행의 속성을 볼 수 있다.

`ws.cell(row=i,column=j).value`: i번째 행 j번째 열의 셀 값을 볼 수 있다.

`flag = True`: 두 개의 테이블을 구분하기 위해서는 테이블 사이를 구분할 방법이 필요하다.

따라서, 셀 값에 원하는 값이 없으면 하나의 테이블이 끝남을 의미하는 것으로 True에서 False로 바꿔준다.

True 상태일 경우는 0번째 리스트에 찾고있는 행 번호를 넣고, False 상태일 경우 1번째 리스트에 넣는다.

0번째 리스트는 첫번째 테이블의 행 값들을 가지고 있고, 1번째 리스트는 두번째 테이블의 행 값들을 가지고 있다.

전체 코드

```python
wb = openpyxl.load_workbook(filename="schedule.xlsx", read_only=False, data_only=True)
ws = wb[wb.sheetnames[0]] # 첫 번째 엑셀시트
visible_rows = [[],[]] # 두 개의 나눠진 테이블을 구하고자 한다
flag = True # 첫 테이블과 두번째 테이블 구분한다

for i in range(4,35):
    if tuple(ws.row_dimensions[i])[0][0] != "hidden": # i번째 행이 숨겨진 행이 아닐 때
        cell_value = ws.cell(row=i,column=2).value # i행 2번째 셀 값
        if cell_value != None: # 빈 칸이 아닐 때
            if "MAERSK" in cell_value or "Blank" in cell_value: # 셀 값에 원하는 값이 있으면
                if flag: # 첫 번째 테이블
                    visible_rows[0].append(i) # 행의 번호를 리스트에 추가
                else: # 두 번째 테이블
                    visible_rows[1].append(i)
            else: # 원하는 값이 없으면 두번째 테이블로 이동
                flag = False
print(visible_rows)

>>>[[10, 11, 12, 13, 14, 15, 16], [26, 27, 28, 29, 30, 31, 32]]
```
