생활 재료의 가격을 거래소에서 가져와 융화 재료의 제작 손익을 표시해주는 Excel 파일입니다.

### 요구 사항
LostArk API 키가 필요합니다. [개발자 포털](https://developer-lostark.game.onstove.com/)에서 API 키를 발급받으세요.

### 사용법 1: XLSM 파일을 사용하는 경우
1. XLSM 파일을 연다.
2. T9 셀의 '제작 수수료 절감률'을 자신의 원정대 영지 환경에 맞춰 수정한다.
3. Alt + F11을 눌러 좌측의 '모듈', 'fetch_market_min_price'를 클릭한다.
   
![image](https://github.com/user-attachments/assets/18d50903-33ce-4e27-9771-e48369078df5)

4. 코드에서 'api_key' 변수에 개발자 포털에서 받은 API 키를 넣는다.

![image](https://github.com/user-attachments/assets/083e29aa-558b-4c35-a4ac-ab6028227f03)

5. 개발자 도구 창을 닫고 '거래소 가격 가져오기' 버튼을 클릭한다.
6. 가격이 바뀌었으면 제작 손익을 확인한다.


### 사용법 2: XLSX 파일을 사용하는 경우
1. XLSX 파일을 사용할 디렉토리에 위치시킨다.
2. main.py 스크립트와 같은 폴더에 '.env' 파일을 생성한다.
3. .env.example의 예시를 따라 개발자 포털에서 받은 API 키와 XLSX 파일의 절대 경로를 입력한다.
4. main.py 스크립트를 실행한다.
5. XLSX 파일의 저장 시간이 바뀌었으면 정상이다. XLSX 파일을 연다.
6. 제작 손익을 확인한다.

* 두 번째 sheet인 record를 사용하여 생활 재료 구매가, 융화 재료 판매가를 기입하면 실제 수익을 확인할 수 있습니다.
