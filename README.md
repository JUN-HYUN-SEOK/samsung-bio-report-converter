# 📊 삼성바이오 레포트 변환기

RAW 데이터를 고객사 요구사항에 맞게 자동 변환하는 Streamlit 웹 애플리케이션입니다.

## 🚀 바로 사용하기

### 온라인 버전 (권장)
👉 **[Streamlit Cloud에서 바로 사용하기](https://your-app-name.streamlit.app)**

### 로컬 실행
```bash
# 1. 저장소 클론
git clone https://github.com/your-username/samsung-bio-report-converter.git
cd samsung-bio-report-converter

# 2. 패키지 설치
pip install -r requirements.txt

# 3. 앱 실행
streamlit run app_streamlit.py
```

## ✨ 주요 기능

- **🔄 자동 매핑**: 244개 컬럼을 미리 정의된 매핑 규칙으로 변환
- **📊 실시간 처리**: 진행률 표시와 함께 빠른 데이터 변환
- **🎨 스타일 적용**: 중요 컬럼 노란색 하이라이트 + 텍스트 형식 지정
- **💾 간편 다운로드**: 브라우저에서 바로 엑셀 파일 다운로드
- **🛡️ 안전한 처리**: 누락 컬럼, NULL 값 등 모든 예외 상황 처리

## 📋 사용 방법

1. **파일 업로드**: RAW 데이터 엑셀 파일 선택
2. **데이터 확인**: 원본 및 변환된 데이터 미리보기
3. **파일 생성**: 스타일이 적용된 엑셀 파일 자동 생성
4. **다운로드**: 변환 완료된 결과 파일 다운로드

## 🔧 기술 스택

- **Frontend**: Streamlit 1.41.1
- **Data Processing**: Pandas 2.2.2
- **Excel Handling**: OpenPyXL 3.1.2
- **Deployment**: Streamlit Cloud

## 📦 파일 구조

```
├── app_streamlit.py    # 메인 Streamlit 앱
├── requirements.txt    # Python 패키지 의존성
├── README.md          # 프로젝트 설명서
└── .streamlit/        # Streamlit 설정 (옵션)
    └── config.toml
```

## 🎯 데이터 변환 로직

### 1. 매핑 딕셔너리
- 고객사 헤더 → RAW 데이터 헤더 매핑
- 총 244개 컬럼 자동 처리

### 2. 데이터 변환
- 숫자 데이터: 3자리 포맷 (001, 002, ...)
- 특별 컬럼: 원본 형식 유지
- NULL 처리: 빈 문자열로 통일

### 3. 스타일 적용
- **노란색 하이라이트**: 신고번호, 정정차수, 신고일자, 수리일자
- **텍스트 형식**: 신고세관, 세관 (숫자 변환 방지)

## 🔍 출력 형식

| 행 번호 | 내용 |
|---------|------|
| 1행 | RAW 데이터 원본 헤더명 |
| 2행 | 변환된 고객사 헤더명 |
| 3행~ | 변환된 실제 데이터 |

## 👨‍💻 개발자 정보

- **개발**: Python 3.12 + Streamlit
- **최적화**: 성능 경고 해결 (pd.concat 사용)
- **안정성**: 전체 예외 처리 및 오류 로깅

## 📞 문의사항

프로그램 사용 중 문제가 발생하면 Issues 탭에 문의해주세요.

---

⭐ **이 프로젝트가 유용하다면 Star를 눌러주세요!** 