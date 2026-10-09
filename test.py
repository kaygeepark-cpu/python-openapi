import json
import os
import requests
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

def fetch_and_parse_data(api_url, file_format="excel", filename="output"):

    api_key = os.getenv("API_KEY")
    # 1. 쿼리 파라미터를 의미 단위의 변수로 분리하여 딕셔너리로 구성
    params = {
        # 기본 요청 설정
        "method": "getList",
        "apiKey": api_key,
        "format": "json",
        "jsonVD": "Y",
        # 조회 대상 및 항목 설정
        "orgId": "354",
        "tblId": "DT_354006_2021A002",
        "itmId": "T001 T002 T003 T004 ",  # 공백은 알아서 인코딩됨
        "outputFields": "ORG_ID ITM_NM LST_CHN_DE ",
        # 계층 구조 파라미터
        "objL1": "ALL",
        "objL2": "",
        "objL3": "",
        "objL4": "",
        "objL5": "",
        "objL6": "",
        "objL7": "",
        "objL8": "",
        # 기간 및 주기 설정
        "prdSe": "Y",
        "newEstPrdCnt": "3",
        "prdInterval": "2025",

        
    }

    try:
        # 2. API 호출 (GET 요청 시 params 전달)
        response = requests.get(api_url, params=params)

        # 요청 성공 여부 확인 (상태 코드 200)
        if response.status_code == 200:
            raw_data = response.json()

            # API 응답 구조에 맞게 리스트 추출 (필요시 경로 수정 가능)
            items = (
                raw_data if isinstance(raw_data, list) else raw_data.get("items", [])
            )

            parsed_data = []
            print("=== 파싱된 데이터 목록 ===")

            # 3. 데이터 예쁘게 파싱하기
            for item in items:
                dt = item.get("DT")
                lst_chn_de = item.get("LST_CHN_DE")
                itm_nm = item.get("ITM_NM")
                org_id = item.get("ORG_ID")

                parsed_item = {
                    "기준일자": lst_chn_de,
                    "항목명": itm_nm,
                    "금액/수치": dt,
                    "기관ID": org_id,
                }
                parsed_data.append(parsed_item)

                print(f"- [{lst_chn_de}] {itm_nm}: {dt} (기관코드: {org_id})")

            # 4. 파일 포맷별 저장 기능 구현
            format_type = file_format.lower()
            if format_type == "excel":
                file_name = f"{filename}.xlsx"
                df = pd.DataFrame(parsed_data)
                df.to_excel(file_name, index=False, engine="openpyxl")
                print(f"엑셀 파일로 저장되었습니다: {file_name}")

            elif format_type == "json":
                file_name = f"{filename}.json"
                with open(file_name, "w", encoding="utf-8") as f:
                    json.dump(
                        parsed_data, f, ensure_ascii=False, indent=4
                    )
                print(f"JSON 파일로 저장되었습니다: {file_name}")

            elif format_type == "txt":
                file_name = f"{filename}.txt"
                with open(file_name, "w", encoding="utf-8") as f:
                    for item in parsed_data:
                        line = f"기준일자: {item['기준일자']}, 항목명: {item['항목명']}, 금액/수치: {item['금액/수치']}, 기관ID: {item['기관ID']}\n"
                        f.write(line)
                print(f"TXT 파일로 저장되었습니다: {file_name}")
            else:
                print(
                    f"지원하지 않는 포맷입니다 ('excel', 'json', 'txt' 중 선택): {file_format}"
                )

            return parsed_data
        else:
            print(f"API 호출 실패: 상태 코드 {response.status_code}")
            print(f"에러 내용: {response.text}")
            return None

    except Exception as e:
        print(f"에러 발생: {e}")
        return None


# 사용 예시
url = "https://kosis.kr/openapi/Param/statisticsParameterData.do"

# 1. 기본값으로 실행 (Excel 저장)
result = fetch_and_parse_data(url)

# 2. JSON 형식으로 저장하고 싶을 때
# result = fetch_and_parse_data(url, file_format="json", filename="my_data")

# 3. TXT 형식으로 저장하고 싶을 때
# result = fetch_and_parse_data(url, file_format="txt", filename="my_data")