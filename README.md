# projectJH

* 파이썬으로 크롤링 하기
* 메인 라이브러리: PyQt5, Selenium
* URL: https://www.maersk.com/schedules
* 목적: 인턴 중인 박주환 도와주기
* 기능: 특정 날짜의 특정 선박의 스케쥴을 크롤링 해오기, 일정 간격으로 자동 업데이트 가능


## ver 1.0 

### 실행 화면

<img width="416" alt="jhresult" src="https://user-images.githubusercontent.com/78152114/141646688-a042a7bb-4721-4ae0-847a-ef6c634035cf.png">

### 사용 설명

<img width="416" alt="jhresult" src="https://user-images.githubusercontent.com/78152114/141647260-27852112-fd22-4485-b3c0-ac3b6a0e616a.png">

Last Updated -2021/11/13 SAT
1. 선박 스케쥴 결과 표시
2. 수동 업데이트
3. 업데이트 정보 표시
4. 자동 업데이트

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
