# MDS 이력 조회 (Streamlit)

기존 로컬 Tkinter(`main_fixed.py`)로 사용하던 MDS 조회 기능을 **Streamlit(웹)** 으로 옮긴 버전입니다.

중요: Tkinter GUI는 Streamlit Cloud(웹)에서 그대로 실행할 수 없어서, **기능/흐름은 최대한 동일하게 유지하면서 웹 UI로 재구성**했습니다.

업로드한 MDS CSV/XLSX 파일을 기준으로, 아래 기능으로 이력을 조회합니다.

- 기간 선택(날짜 범위)
- 결함 내용 검색
- 조치 내용 검색
- 기종 선택(멀티)
- 기번(시리얼) 검색
- 컬럼 표시 선택
- 필터 결과 CSV 다운로드
- AC Type / Reg / W/O / ATA 등 기존 검색조건
- 반복결함(정비규정 4.3.3.10) 자동 분석
- 통계 분석(ATA 챕터별)
- Gemini 2.0 채팅(옵션)

## 로컬 실행

```bash
cd tway_mds_streamlit
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Community Cloud 배포

1. 이 폴더(`tway_mds_streamlit/`)를 GitHub 공개 레포로 올립니다.
2. Streamlit Community Cloud에서 New app → 레포 선택
3. Main file path에 `tway_mds_streamlit/app.py` 지정
4. Deploy

## Gemini API Key (옵션)

- 로컬 실행: 사이드바에 입력해서 사용 가능
- Streamlit Cloud: `Settings → Secrets`에 아래처럼 넣으면 안전합니다.

```toml
GEMINI_API_KEY = "YOUR_KEY"
```

## 기본 컬럼 매핑

기본값(필요하면 사이드바에서 변경 가능):

- 기간(날짜): `Noti. Date`
- 결함 내용: `Noti. Desc.`
- 조치 내용: `Corrective Action`
- 기종: `Equip.`
- 기번(시리얼): `MFG S/N`
