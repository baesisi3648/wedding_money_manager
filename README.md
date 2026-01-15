# 👰 Wedding Money Manager (축의금 관리 도우미)

인터넷 연결이 없는 결혼식장 접수대 환경을 고려하여 개발된 **Windows 기반 축의금 관리 및 정산 프로그램**입니다.
Python과 PyQt6로 개발되었으며, 복잡한 현장 상황에서도 빠르고 정확하게 축의금 내역을 기록하고 엑셀로 정산할 수 있습니다.

## 📥 다운로드 및 실행 (Download)
**개발 환경 설정 없이 바로 실행 가능한 프로그램(.exe)을 다운로드하세요.**

1. 👉 **[최신 버전 다운로드 페이지 (GitHub Releases)](https://github.com/baesisi3648/wedding_money_manager/releases)** 로 이동합니다.
2. `Assets` 항목에 있는 **`축의금관리_v1.0.exe`** 파일을 클릭하여 다운로드합니다.
3. 별도의 설치 과정 없이, 다운로드한 파일을 더블 클릭하면 바로 실행됩니다.
   > *주의: 실행 시 Windows 보안 경고가 뜰 경우, '추가 정보' -> '실행'을 클릭하시면 됩니다.*

---

## 📸 실행 화면
*<img width="1379" height="1032" alt="image" src="https://github.com/user-attachments/assets/c68b0a69-ca13-4468-b82f-a054602c7664" />*

## ✨ 주요 기능
- **오프라인 구동:** 로컬 DB(SQLite)를 사용하여 인터넷 없이도 완벽하게 작동
- **실시간 대시보드:** 총 인원, 축의금 합계, 식권 배부 현황 실시간 집계
- **빠른 입력 UX:** 키보드(Tab/Enter) 만으로 입력 가능한 최적화된 동선 제공
- **자동 포맷팅:** 금액 입력 시 3자리 콤마(,) 자동 적용 및 봉투 번호 자동 부여
- **안전한 데이터 관리:**
  - 프로그램 강제 종료 시에도 데이터가 유지되는 SQLite 트랜잭션 처리
  - 동명이인 등록 시 중복 경고 알림
- **엑셀 정산:** 접수 마감 후 `OpenPyXL`을 활용한 상세/요약 보고서 자동 생성 (`.xlsx`)
- **다크 모드 대응:** 사용자 윈도우 테마와 관계없이 일관된 가시성을 제공하는 UI 스타일링

## 🛠 기술 스택 (Tech Stack)
- **Language:** Python 3.13
- **GUI:** PyQt6
- **Database:** SQLite3
- **Data Processing:** OpenPyXL (Excel Export)
- **Distribution:** PyInstaller (Standalone .exe)

## 🚀 문제 해결 및 성능 최적화 (Optimization Story)
### 1. 용량 경량화 (400MB → 45MB)
- 초기 개발 단계에서 `Pandas` 라이브러리를 사용했으나, 단순 엑셀 내보내기 기능 대비 실행 파일 용량이 과도하게 커지는 문제(약 400MB) 발생.
- 무거운 `Pandas` 의존성을 제거하고, 가볍고 효율적인 `OpenPyXL`로 코드를 리팩토링함.
- `venv` 가상환경에서 불필요한 라이브러리를 배제한 Clean Build를 수행하여 최종 결과물 용량을 **약 89% 감소**시킴.

### 2. 다크 모드 가시성 문제 해결
- 일부 사용자의 Windows 다크 모드 환경에서 테이블 내 텍스트가 흰색 배경에 흰색 글씨로 출력되어 보이지 않는 UI 버그 발견.
- PyQt의 `setStyleSheet`를 활용하여 OS 테마 설정과 무관하게 텍스트 및 배경 색상을 강제 고정(Hard-coding)하여 일관된 UX를 확보함.

## 📂 프로젝트 구조
```bash
Wedding-Money-Manager/
├── wedding_manager.py   # 메인 소스 코드
├── requirements.txt     # 의존성 라이브러리 목록
└── README.md            # 프로젝트 문서
