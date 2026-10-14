from abc import ABC, abstractmethod
from dataclasses import dataclass
import datetime
import json
import os
import re
import time
import urllib.parse
import FinanceDataReader as fdr
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import requests
from bs4 import BeautifulSoup

# 구글 시트 및 텔레그램 전송용 라이브러리 (깃허브 액션 환경에서 자동 설치됨)
try:
    import gspread
    from google.oauth2.service_account import Credentials

    GOOGLE_API_AVAILABLE = True
except ImportError:
    GOOGLE_API_AVAILABLE = False


# ==========================================
# 1. 표준 데이터 모델 & 증분 수집기
# ==========================================
@dataclass
class ContentItem:
    source_type: str
    channel_name: str
    content_id: str
    title: str
    published_date: datetime.date
    text: str
    url: str


class BaseCollector(ABC):

    @abstractmethod
    def collect(
        self, days_limit: int, known_ids: set[str]
    ) -> list[ContentItem]:
        pass


class NaverBlogCollector(BaseCollector):
    """네이버 블로그 전용 수집기 (기사 링크 박스 및 URL 텍스트 자동 제거)"""

    def __init__(self, blog_id: str, blog_name: str):
        self.blog_id = blog_id
        self.blog_name = blog_name
        self.headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
            )
        }

    def collect(
        self, days_limit: int = 365, known_ids: set[str] = None
    ) -> list[ContentItem]:
        if known_ids is None:
            known_ids = set()

        items = []
        cutoff_date = datetime.date.today() - datetime.timedelta(
            days=days_limit
        )

        if known_ids:
            print(
                f"▶ [{self.blog_name}] 기존 장부 확인 완료! 새로 올라온 글만 빠르게 탐색합니다..."
            )
        else:
            print(
                f"▶ [{self.blog_name}] 최근 {days_limit}일 치 글을 수집합니다 (기준일: {cutoff_date} 이후)..."
            )

        page = 1
        post_metas = []
        while True:
            list_url = (
                "https://blog.naver.com/PostTitleListAsync.naver?"
                f"blogId={self.blog_id}&currentPage={page}&categoryNo=0&countPerPage=30"
            )
            res = requests.get(list_url, headers=self.headers)
            if res.status_code != 200:
                break

            raw_text = res.text.replace("\\'", "'")
            try:
                data = requests.models.complexjson.loads(raw_text)
            except Exception:
                break

            post_list = data.get("postList", [])
            if not post_list:
                break

            stop_flag = False
            for post in post_list:
                log_no = str(post.get("logNo")).strip()
                if log_no in known_ids:
                    stop_flag = True
                    break

                raw_title = urllib.parse.unquote(post.get("title", "")).replace(
                    "+", " "
                )
                add_date_str = (
                    post.get("addDate", "").replace(" ", "").rstrip(".")
                )

                try:
                    post_date = datetime.datetime.strptime(
                        add_date_str, "%Y.%m.%d"
                    ).date()
                except Exception:
                    continue

                if post_date < cutoff_date:
                    stop_flag = True
                    break

                post_metas.append((log_no, raw_title, post_date))

            if stop_flag:
                break
            page += 1
            time.sleep(0.2)

        if not post_metas:
            print(f"▶ [{self.blog_name}] 새로 추가된 글이 없습니다.")
            return []

        print(
            f"▶ [{self.blog_name}] 신규 글 {len(post_metas)}개를 찾았습니다. 본문을 읽는 중..."
        )
        for idx, (log_no, title, post_date) in enumerate(post_metas, 1):
            text = self._get_post_body(log_no)
            if text:
                items.append(
                    ContentItem(
                        source_type="NaverBlog",
                        channel_name=self.blog_name,
                        content_id=log_no,
                        title=title,
                        published_date=post_date,
                        text=f"{title} . {text}",
                        url=f"https://blog.naver.com/{self.blog_id}/{log_no}",
                    )
                )
            if idx % 5 == 0 or idx == len(post_metas):
                print(f"   ... {idx}/{len(post_metas)}개 신규 글 읽기 완료")
            time.sleep(0.35)

        return items

    def _get_post_body(self, log_no: str) -> str:
        url = f"https://blog.naver.com/PostView.naver?blogId={self.blog_id}&logNo={log_no}"
        res = requests.get(url, headers=self.headers)
        if res.status_code != 200:
            return ""
        soup = BeautifulSoup(res.text, "html.parser")
        container = soup.find("div", class_="se-main-container") or soup.find(
            "div", id="postViewArea"
        )
        if not container:
            return ""

        for oglink in container.find_all("div", class_="se-module-oglink"):
            oglink.decompose()

        raw_text = container.get_text(separator=" ", strip=True)
        clean_text = re.sub(r"https?://\S+|www\.\S+", " ", raw_text)
        return clean_text


# ==========================================
# 2. 종목 사전 및 주가 흐름 분석 엔진
# ==========================================
class StockEngine:

    def __init__(self):
        self.stock_map = {}
        self.price_cache = {}
        self.negative_words = [
            "위험",
            "유의",
            "주의",
            "조심",
            "피해",
            "피하",
            "하락",
            "폭락",
            "급락",
            "적자",
            "부실",
            "부도",
            "파산",
            "워크아웃",
            "법정관리",
            "상장폐지",
            "상폐",
            "거품",
            "고평가",
            "과열",
            "악재",
            "우려",
            "타격",
            "손실",
            "차익실현",
        ]

        self.custom_theme_dict = {
            "삼성전자": "🧠AI반도체/HBM·메모리 및 스마트폰·가전 세계 1위",
            "SK하이닉스": "🧠AI반도체/HBM(고대역폭메모리) 글로벌 1위 핵심 수혜주",
            "한미반도체": "🧠AI반도체/HBM 필수 제조장비(TC본더) 글로벌 독점적 기업",
            "리노공업": "🧠반도체 검사용 소켓·핀 세계적 부품사",
            "이오테크닉스": "🧠반도체 레이저 장비 주력",
            "주성엔지니어링": "🧠반도체/디스플레이 증착 장비",
            "NAVER": "🌐국내 최대 포털·쇼핑 및 하이퍼클로바X AI 플랫폼",
            "카카오": "💬국내 최대 메신저·모빌리티·핀테크 플랫폼",
            "HD현대일렉트릭": "⚡전력기기/초고압변압기 (미국 AI 데이터센터 전력망 확충 수혜)",
            "효성중공업": "⚡전력기기/초고압변압기 및 차단기, 수소 충전소",
            "제룡전기": "⚡전력기기/배전변압기 (미국 배전망 교체 집중 수혜)",
            "LS일렉트릭": "⚡전력기기/배전반·전력시스템 및 초고압변압기",
            "LS": "⚡전선/구리/해저케이블 및 전력 인프라 지주사",
            "대한전선": "⚡전선/초고압케이블 및 해저케이블(구리·전력망 테마)",
            "가온전선": "⚡전선/전력케이블 제조 (LS그룹 계열)",
            "일진전기": "⚡전력기기/초고압전선 및 변압기 동시 영위",
            "산일전기": "⚡전력기기/특수변압기 (신재생·데이터센터·전력망용)",
            "두산에너빌리티": "☢️원전/SMR(소형모듈원전) 주기기 제작 및 가스터빈 대표주",
            "한전기술": "☢️원전 설계 및 엔지니어링 독점적 기업",
            "한전KPS": "☢️원전 및 발전소 정비·유지보수 전문",
            "비에이치아이": "☢️발전용 보일러·HRSG(복합화력) 및 원전 보조기기",
            "우진": "☢️원전용 계측기 핵심 부품사",
            "한화에어로스페이스": "🛡️방산/우주항공 (K9 자주포, 천무, 항공엔진 대장주)",
            "현대로템": "🛡️방산/철도 (K2 전차 폴란드·유럽 수출 및 고속철)",
            "LIG넥스원": "🛡️방산/정밀유도무기 (천궁-II 미사일 수출 주역)",
            "한국항공우주": "🛡️방산/우주항공 (KF-21 전투기, FA-50 경공격기 제작)",
            "한화시스템": "🛡️방산 전자장비(레이다) 및 저궤도 위성통신",
            "풍산": "🛡️방산(포탄·탄약 수출) 및 구리(신동) 가격 상승 동시 수혜",
            "HD한국조선해양": "🚢조선 중간지주사 (LNG선·암모니아선·미 해군 MRO)",
            "HD현대중공업": "🚢조선/해양방산 (글로벌 1위 조선소, 잠수함·함정)",
            "한화오션": "🚢조선/특수선 (LNG운반선, 잠수함 및 미 해군 함정 MRO 수혜)",
            "삼성중공업": "🚢조선/FLNG(부유식 액화천연가스 생산설비) 글로벌 강자",
            "HD현대마린솔루션": "🚢선박 애프터마켓(AM) 정비·부품 및 친환경 개조",
            "HMM": "🚢해운 (컨테이너선 운임 상승 및 홍해 사태 수혜)",
            "팬오션": "🚢해운 (벌크선·곡물·원자재 운송 및 LNG선)",
            "삼양식품": "🍜K-푸드/수출 (불닭볶음면 글로벌 열풍 대장주)",
            "농심": "🍜K-푸드/라면 글로벌 수출 확대",
            "실리콘투": "💄K-뷰티/화장품 글로벌 역직구 유통 플랫폼",
            "에이피알": "💄K-뷰티/홈뷰티 디바이스(메디큐브) 글로벌 수출",
            "한국콜마": "💄K-뷰티/화장품 ODM(위탁생산) 대장주",
            "코스맥스": "💄K-뷰티/글로벌 1위 화장품 ODM 기업",
            "삼성바이오로직스": "💊바이오/CDMO(항체의약품 위탁생산) 글로벌 1위 (생물보안법 수혜)",
            "셀트리온": "💊바이오시밀러 및 신약 개발 대표주",
            "유한양행": "💊제약/폐암 신약 '렉라자' FDA 승인 및 기술수출",
            "알테오젠": "💊바이오 플랫폼 (정맥주사를 피하주사로 바꾸는 SC 제형 기술수출)",
            "리가켐바이오": "💊차세대 항암제 ADC(항체-약물 접합체) 플랫폼 강자",
            "고려아연": "⛏️비철금속 제련 세계 1위 (금·은·아연·희소금속 및 경영권 이슈)",
            "포스코홀딩스": "🔋철강 및 2차전지 소재(리튬·니켈) 지주사",
            "LG에너지솔루션": "🔋2차전지/전기차 배터리 및 ESS 국내 1위",
            "삼성SDI": "🔋2차전지/각형 배터리 및 전고체 배터리",
            "에코프로비엠": "🔋2차전지 하이니켈 양극재 대표 기업",
            "현대차": "🚗완성차/하이브리드·전기차 및 인도법인·로보틱스",
            "기아": "🚗완성차/RV·하이브리드 고수익성 및 전기차 라인업",
            "KB금융": "🏦금융/은행 대장주 (기업 밸류업·주주환원·배당 대표)",
            "신한지주": "🏦금융/은행 (밸류업 및 자사주 소각·배당)",
            "메리츠금융지주": "🏦금융/자사주 매입·소각 주주환원 끝판왕",
        }

        self.issue_keywords = [
            "AI",
            "데이터센터",
            "HBM",
            "반도체",
            "엔비디아",
            "전력망",
            "변압기",
            "전선",
            "구리",
            "원전",
            "SMR",
            "우라늄",
            "가스터빈",
            "신재생",
            "태양광",
            "풍력",
            "ESS",
            "방산",
            "우주",
            "위성",
            "미사일",
            "전차",
            "자주포",
            "탄약",
            "전쟁",
            "중동",
            "조선",
            "LNG",
            "해운",
            "운임",
            "MRO",
            "함정",
            "잠수함",
            "쇄빙선",
            "비만치료제",
            "바이오",
            "CDMO",
            "생물보안법",
            "신약",
            "FDA",
            "임상",
            "ADC",
            "트럼프",
            "해리스",
            "관세",
            "미중갈등",
            "공급망",
            "보조금",
            "리쇼어링",
            "금리인하",
            "금리",
            "연준",
            "환율",
            "엔화",
            "밸류업",
            "배당",
            "자사주",
            "수주",
            "실적",
            "어닝서프라이즈",
            "수출",
            "품귀",
            "공급부족",
            "독점",
            "양자컴퓨터",
            "로봇",
            "자율주행",
            "전기차",
            "하이브리드",
            "2차전지",
            "리튬",
            "원유",
            "천연가스",
            "금값",
            "은값",
            "희토류",
            "곡물",
            "불닭",
            "화장품",
            "중국",
            "부양책",
            "이구환신",
        ]

        self._build_universe()
        self.pattern = self._compile_regex()

    def _build_universe(self):
        exclude_kr = {
            "NAVER",
            "SBS",
            "YTN",
            "KNN",
            "IHQ",
            "CJ",
            "SK",
            "LG",
            "GS",
            "LS",
            "캐리",
            "남성",
            "씨앗",
            "배럴",
            "대상",
            "모두투어",
            "위메이드",
            "원익",
            "선도전기",
            "이수",
            "선익",
            "신영",
            "대신",
            "한양",
            "서울",
            "부산",
            "동양",
            "대한",
            "제일",
            "고려",
            "우리",
            "한창",
            "미래",
            "흥국",
            "보령",
            "성안",
            "대원",
            "삼영",
            "동원",
            "태광",
            "세방",
            "화성",
            "우진",
            "신성",
            "대유",
            "삼천리",
            "국보",
            "백광",
            "조일",
            "남선",
            "동방",
            "전방",
            "방림",
            "일신",
            "대동",
            "대창",
            "동국",
            "동서",
            "동신",
            "동아",
            "동일",
            "삼보",
            "삼진",
            "삼화",
            "성신",
            "세기",
            "세원",
            "세진",
            "신라",
            "신원",
            "신일",
            "신흥",
            "아세아",
            "영풍",
            "유니온",
            "유성",
            "유진",
            "일성",
            "일양",
            "일정",
            "일진",
            "중앙",
            "진흥",
            "태양",
            "태영",
            "태원",
            "평화",
            "한국",
            "한성",
            "한신",
            "한일",
            "현대",
            "화신",
            "화천",
            "환인",
            "흥아",
            "가온",
            "나무",
            "나래",
            "누리",
            "다산",
            "대림",
            "대명",
            "대성",
            "대주",
            "대호",
            "덕산",
            "덕성",
            "덕양",
            "도화",
            "동구",
            "동성",
            "동우",
            "동운",
            "동화",
            "두산",
            "디아이",
            "레이",
            "리노",
            "마니커",
            "모베이스",
            "무학",
            "바다",
            "보광",
            "보해",
            "부국",
            "비츠로",
            "사조",
            "삼기",
            "삼륭",
            "삼목",
            "삼부",
            "상상인",
            "상신",
            "새론",
            "서연",
            "서원",
            "서진",
            "서희",
            "선진",
            "성광",
            "성도",
            "성문",
            "성보",
            "성우",
            "성호",
            "세명",
            "세보",
            "세아",
            "세우",
            "세운",
            "세코닉스",
            "세한",
            "소프트센",
            "솔본",
            "수산",
            "시공테크",
            "시너지",
            "신풍",
            "쌍용",
            "아이콘",
            "아진",
            "안국",
            "알파",
            "에스엘",
            "에이스",
            "에이프로",
            "엔케이",
            "엘앤씨",
            "엠케이",
            "영보",
            "영신",
            "영원",
            "영진",
            "영흥",
            "오리온",
            "오성",
            "오스코텍",
            "와이지",
            "우림",
            "우성",
            "우신",
            "우원",
            "우정",
            "원풍",
            "유니드",
            "유니슨",
            "유니크",
            "유라",
            "유비쿼스",
            "유신",
            "유안타",
            "유유",
            "유화",
            "율촌",
            "이구산업",
            "이니텍",
            "이노션",
            "이더블유",
            "이랜텍",
            "이리츠",
            "이마트",
            "이연",
            "이지",
            "이화",
            "인디에프",
            "인산가",
            "인성",
            "인스코비",
            "인탑스",
            "인포바인",
            "인피니트",
            "일동",
            "일야",
            "일지테크",
            "자화전자",
            "장원",
            "재영",
            "정다운",
            "정상제이엘에스",
            "정산애강",
            "정원",
            "제노포커스",
            "제우스",
            "제이준",
            "제일연마",
            "제주",
            "조광",
            "조비",
            "조선",
            "조아",
            "종근당",
            "좋은사람들",
            "주성",
            "주연",
            "삼양",
            "금양",
            "금호",
            "한화",
            "효성",
            "코오롱",
            "웅진",
            "교보",
            "다우",
        }

        sector_map_kr = {}
        try:
            df_desc = fdr.StockListing("KRX-DESC")
            for _, r in df_desc.iterrows():
                c = str(r.get("Code", "")).strip()
                sec = str(r.get("Sector", "")).strip()
                ind = str(r.get("Industry", "")).strip()
                if c:
                    sector_map_kr[c] = (
                        f"{sec} ({ind[:25]})"
                        if sec and sec != "nan" and ind and ind != "nan"
                        else (sec if sec != "nan" else ind)
                    )
        except Exception:
            pass

        df_krx = fdr.StockListing("KRX")
        for _, row in df_krx.iterrows():
            name = str(row["Name"]).strip()
            code = str(row["Code"]).strip()
            if "스팩" in name or name in exclude_kr:
                continue
            if len(name) == 2:
                marcap = row.get("Marcap", 0)
                if pd.notna(marcap) and float(marcap) < 300_000_000_000:
                    continue

            raw_sector = (
                sector_map_kr.get(code, "")
                or str(row.get("Sector", "")).strip()
                or str(row.get("Dept", "")).strip()
            )
            if raw_sector == "nan" or not raw_sector:
                raw_sector = "국내상장사"
            theme_desc = self.custom_theme_dict.get(name, raw_sector)

            if len(name) >= 2:
                self.stock_map[name] = {
                    "std_name": name,
                    "code": code,
                    "market": "KR",
                    "sector": raw_sector,
                    "theme_desc": theme_desc,
                }

        try:
            df_sp500 = fdr.StockListing("S&P500")
            exclude_us = {
                "ALL",
                "CAN",
                "SEE",
                "ARE",
                "NOW",
                "ONE",
                "OUT",
                "HAS",
                "KEY",
                "LOW",
                "DAY",
                "NEW",
                "BIG",
                "TOP",
                "FOR",
                "AND",
                "THE",
                "YOU",
                "NOT",
                "GET",
                "MAY",
                "WAY",
                "BRO",
                "CAT",
                "COP",
                "COST",
                "FAST",
                "FIT",
                "FIVE",
                "FOX",
                "FUN",
                "GAP",
                "GEN",
                "GOLD",
                "GOOD",
                "HAS",
                "HES",
                "HIGH",
                "HIT",
                "HOME",
                "HON",
                "HOPE",
                "HOT",
                "ICE",
                "ING",
                "ITT",
                "JOB",
                "JOY",
                "KEY",
                "KIM",
                "LAW",
                "LEG",
                "LEN",
                "LIFE",
                "LINE",
                "LIVE",
                "LOVE",
                "LOW",
                "MAN",
                "MAP",
                "MAR",
                "MAS",
                "MAT",
                "MAX",
                "MED",
                "MEN",
                "MET",
                "MID",
                "MIN",
                "MOM",
                "NET",
                "NEXT",
                "NICE",
                "NINE",
                "NOW",
                "OIL",
                "OLD",
                "ONE",
                "OPEN",
                "OUR",
                "OUT",
                "PAY",
                "PEG",
                "PEP",
                "PET",
                "PLAY",
                "PLUS",
                "POOL",
                "POST",
                "PRO",
                "PUT",
                "PWR",
                "RAW",
                "RAY",
                "RED",
                "REG",
                "RENT",
                "RES",
                "RICH",
                "RIDE",
                "RING",
                "ROE",
                "ROL",
                "ROPE",
                "ROSE",
                "RUN",
                "SAFE",
                "SAVE",
                "SAY",
                "SEA",
                "SEE",
                "SELF",
                "SELL",
                "SET",
                "SHIP",
                "SHOP",
                "SHOW",
                "SITE",
                "SIX",
                "SKY",
                "SLOW",
                "SNOW",
                "SOON",
                "SPOT",
                "STAR",
                "STAY",
                "STEP",
                "STOP",
                "SUN",
                "SURE",
                "TALK",
                "TAX",
                "TEAM",
                "TECH",
                "TELL",
                "TEN",
                "TERM",
                "TEST",
                "TEXT",
                "TIME",
                "TIN",
                "TIP",
                "TRUE",
                "TRY",
                "TURN",
                "TWO",
                "TYPE",
                "UNIT",
                "USE",
                "VAN",
                "VERY",
                "VIEW",
                "VOTE",
                "WAIT",
                "WALK",
                "WALL",
                "WANT",
                "WAR",
                "WASH",
                "WAVE",
                "WAY",
                "WEAR",
                "WEEK",
                "WELL",
                "WEST",
                "WHAT",
                "WHEN",
                "WHO",
                "WIDE",
                "WIFI",
                "WILL",
                "WIN",
                "WIND",
                "WINE",
                "WING",
                "WIRE",
                "WISE",
                "WISH",
                "WITH",
                "WOOD",
                "WORD",
                "WORK",
                "YEAR",
                "YES",
                "YET",
                "YOU",
                "ZERO",
                "ZONE",
                "COO",
                "CEO",
                "CFO",
                "CTO",
            }
            for _, row in df_sp500.iterrows():
                sym = str(row["Symbol"]).strip()
                sec = str(row.get("Sector", "미국주식")).strip()
                us_name = str(row.get("Name", sym)).strip()
                if len(sym) >= 3 and sym not in exclude_us:
                    self.stock_map[sym] = {
                        "std_name": sym,
                        "code": sym,
                        "market": "US",
                        "sector": sec,
                        "theme_desc": f"🇺🇸 {us_name} ({sec})",
                    }
        except Exception:
            pass

        us_aliases = {
            "엔비디아": (
                "엔비디아",
                "NVDA",
                "반도체/AI",
                "🧠AI반도체/GPU 독점 세계 1위 (데이터센터 핵심)",
            ),
            "테슬라": (
                "테슬라",
                "TSLA",
                "전기차/자율주행",
                "🚗전기차·자율주행(FSD)·휴머노이드 로봇·ESS 대장주",
            ),
            "애플": (
                "애플",
                "AAPL",
                "빅테크/IT",
                "📱아이폰 및 온디바이스 AI 생태계",
            ),
            "마이크로소프트": (
                "마이크로소프트",
                "MSFT",
                "빅테크/클라우드",
                "☁️클라우드(Azure) 및 오픈AI(ChatGPT) 최대주주",
            ),
            "알파벳": (
                "알파벳",
                "GOOGL",
                "빅테크/AI",
                "🔍구글 검색·유튜브·클라우드 및 제미나이 AI",
            ),
            "구글": (
                "알파벳",
                "GOOGL",
                "빅테크/AI",
                "🔍구글 검색·유튜브·클라우드 및 제미나이 AI",
            ),
            "아마존": (
                "아마존",
                "AMZN",
                "클라우드/이커머스",
                "☁️글로벌 클라우드(AWS) 1위 및 북미 최대 전자상거래",
            ),
            "메타": (
                "메타",
                "META",
                "빅테크/SNS",
                "📱인스타그램·페이스북 및 오픈소스 AI(라마) 선두",
            ),
            "브로드컴": (
                "브로드컴",
                "AVGO",
                "반도체/통신",
                "🧠빅테크 맞춤형 AI칩(ASIC) 및 데이터센터 네트워크 반도체 1위",
            ),
            "TSMC": (
                "TSMC",
                "TSM",
                "반도체 파운드리",
                "🧠글로벌 반도체 위탁생산(파운드리) 독점적 1위 (대만/미국ADR)",
            ),
            "ASML": (
                "ASML",
                "ASML",
                "반도체 장비",
                "🧠초미세 공정 필수 EUV 노광장비 전 세계 유일 독점 기업",
            ),
            "인텔": (
                "인텔",
                "INTC",
                "반도체",
                "🧠PC·서버용 CPU 및 미국 반도체 제조 부활(파운드리) 지원 대상",
            ),
            "AMD": (
                "AMD",
                "AMD",
                "반도체/AI",
                "🧠엔비디아 대항마 AI 가속기(MI300) 및 서버용 CPU 강자",
            ),
            "퀄컴": (
                "퀄컴",
                "QCOM",
                "반도체/통신",
                "📱스마트폰/AI PC용 스냅드래곤 칩셋 강자",
            ),
            "보잉": (
                "보잉",
                "BA",
                "항공/방산",
                "✈️글로벌 양대 민항기·방산·우주 제조사",
            ),
            "록히드마틴": (
                "록히드마틴",
                "LMT",
                "방산/우주",
                "🛡️세계 1위 방위산업체 (F-35 스텔스기, 미사일 방어체계)",
            ),
            "일라이릴리": (
                "일라이릴리",
                "LLY",
                "제약/바이오",
                "💊글로벌 시총 1위 제약사 (비만치료제 '젭바운드/마운자로' 열풍)",
            ),
            "노보노디스크": (
                "노보노디스크",
                "NVO",
                "제약/바이오",
                "💊비만치료제 '위고비'·당뇨약 '오젬픽' 개발사",
            ),
            "엑슨모빌": (
                "엑슨모빌",
                "XOM",
                "에너지/정유",
                "🛢️미국 최대 석유·천연가스 메이저 기업",
            ),
            "셰브론": (
                "셰브론",
                "CVX",
                "에너지/정유",
                "🛢️글로벌 석유·천연가스 메이저 기업",
            ),
            "버크셔해서웨이": (
                "버크셔해서웨이",
                "BRK-B",
                "금융/지주",
                "🏦워런 버핏의 투자 지주사 (보험·철도·에너지·막대한 현금보유)",
            ),
            "팔란티어": (
                "팔란티어",
                "PLTR",
                "AI 소프트웨어",
                "🛡️미국 국방부·정보기관 및 기업용 AI 빅데이터 분석 플랫폼 대장주",
            ),
            "마이크론": (
                "마이크론",
                "MU",
                "반도체/메모리",
                "🧠미국 유일 메모리 반도체(DRAM·HBM) 기업 (삼성·하이닉스 경쟁사)",
            ),
            "버티브": (
                "버티브",
                "VRT",
                "전력/데이터센터",
                "⚡AI 데이터센터 열관리(액체냉각) 및 전력설비 핵심 수혜주",
            ),
            "이튼": (
                "이튼",
                "ETN",
                "전력기기",
                "⚡미국 전력망 확충 및 데이터센터 전력관리 솔루션 대장주",
            ),
            "GE버노바": (
                "GE버노바",
                "GEV",
                "전력/에너지",
                "⚡GE에서 분사한 가스터빈·전력망·풍력·SMR 원전 종합 에너지 기업",
            ),
            "뉴스케일파워": (
                "뉴스케일파워",
                "SMR",
                "원전/SMR",
                "☢️미국 대표 소형모듈원전(SMR) 설계 기업 (두산에너빌리티 파트너)",
            ),
            "오클로": (
                "오클로",
                "OKLO",
                "원전/SMR",
                "☢️샘 올트먼(오픈AI)이 투자한 차세대 소형원전(SMR) 기업",
            ),
            "아이온큐": (
                "아이온큐",
                "IONQ",
                "양자컴퓨터",
                "💻이온트랩 방식 양자컴퓨터 상용화 선두 기업",
            ),
            "코히런트": (
                "코히런트",
                "COHR",
                "광통신/레이저",
                "🌐AI 데이터센터용 초고속 광트랜시버(광통신 부품) 핵심 기업",
            ),
        }
        for alias, (std_name, ticker, sec, desc) in us_aliases.items():
            self.stock_map[alias] = {
                "std_name": std_name,
                "code": ticker,
                "market": "US",
                "sector": sec,
                "theme_desc": desc,
            }

        cn_aliases = {
            "비야디": (
                "비야디(BYD)",
                "1211.HK",
                "전기차/배터리(홍콩HKD)",
                "🇨🇳중국 전기차 및 리튬인산철(LFP) 배터리 세계 1위 기업",
            ),
            "BYD": (
                "비야디(BYD)",
                "1211.HK",
                "전기차/배터리(홍콩HKD)",
                "🇨🇳중국 전기차 및 리튬인산철(LFP) 배터리 세계 1위 기업",
            ),
            "CATL": (
                "CATL(닝더스다이)",
                "300750.SZ",
                "2차전지(선전CNY)",
                "🇨🇳글로벌 전기차 배터리 점유율 독보적 1위 기업",
            ),
            "닝더스다이": (
                "CATL(닝더스다이)",
                "300750.SZ",
                "2차전지(선전CNY)",
                "🇨🇳글로벌 전기차 배터리 점유율 독보적 1위 기업",
            ),
            "알리바바": (
                "알리바바",
                "BABA",
                "이커머스/클라우드(미국ADR)",
                "🇨🇳중국 최대 전자상거래(타오바오·알리익스프레스) 및 클라우드 빅테크",
            ),
            "텐센트": (
                "텐센트",
                "0700.HK",
                "IT/게임/메신저(홍콩HKD)",
                "🇨🇳중국 시총 1위 빅테크 (위챗 메신저 및 글로벌 최대 게임사)",
            ),
            "샤오미": (
                "샤오미",
                "1810.HK",
                "스마트폰/전기차(홍콩HKD)",
                "🇨🇳가성비 스마트폰·가전 및 전기차(SU7) 흥행 대장주",
            ),
            "SMIC": (
                "SMIC(중신궈지)",
                "0981.HK",
                "반도체 파운드리(홍콩HKD)",
                "🇨🇳중국 최대 반도체 위탁생산(파운드리) 국책 핵심 기업",
            ),
            "중신궈지": (
                "SMIC(중신궈지)",
                "0981.HK",
                "반도체 파운드리(홍콩HKD)",
                "🇨🇳중국 최대 반도체 위탁생산(파운드리) 국책 핵심 기업",
            ),
            "우시바이오로직스": (
                "우시바이오로직스",
                "2269.HK",
                "바이오CDMO(홍콩HKD)",
                "🇨🇳중국 대표 바이오 의약품 위탁생산 기업 (미국 생물보안법 이슈 핵심)",
            ),
            "우시바이오": (
                "우시바이오로직스",
                "2269.HK",
                "바이오CDMO(홍콩HKD)",
                "🇨🇳중국 대표 바이오 의약품 위탁생산 기업 (미국 생물보안법 이슈 핵심)",
            ),
            "바이두": (
                "바이두",
                "BIDU",
                "AI/검색엔진(미국ADR)",
                "🇨🇳중국판 구글 (검색엔진 1위 및 자율주행 로보택시·AI 선두)",
            ),
            "핀둬둬": (
                "핀둬둬(테무)",
                "PDD",
                "이커머스(미국ADR)",
                "🇨🇳초저가 직구 앱 '테무(Temu)' 모회사",
            ),
            "테무": (
                "핀둬둬(테무)",
                "PDD",
                "이커머스(미국ADR)",
                "🇨🇳초저가 직구 앱 '테무(Temu)' 모회사",
            ),
            "징동닷컴": (
                "징동닷컴",
                "JD",
                "이커머스/물류(미국ADR)",
                "🇨🇳중국 최대 직매입 기반 전자상거래 및 물류 기업",
            ),
            "메이투안": (
                "메이투안",
                "3690.HK",
                "배달/로컬플랫폼(홍콩HKD)",
                "🇨🇳중국 배달의민족 (음식배달·로컬서비스 독점적 1위)",
            ),
            "리오토": (
                "리오토(이상자동차)",
                "LI",
                "전기차(미국ADR)",
                "🇨🇳중국 주행거리연장형(EREV) 프리미엄 전기차 선두주자",
            ),
            "샤오펑": (
                "샤오펑",
                "XPEV",
                "전기차/자율주행(미국ADR)",
                "🇨🇳중국 자율주행 특화 전기차 스타트업",
            ),
            "니오": (
                "니오",
                "NIO",
                "전기차(미국ADR)",
                "🇨🇳배터리 교체형 중국 프리미엄 전기차 기업",
            ),
            "지리자동차": (
                "지리자동차",
                "0175.HK",
                "완성차/전기차(홍콩HKD)",
                "🇨🇳볼보 모회사이자 중국 대표 민영 완성차 기업",
            ),
            "귀주모태주": (
                "귀주모태주(마오타이)",
                "600519.SS",
                "필수소비재/주류(상하이CNY)",
                "🇨🇳중국 최고급 백주 '마오타이' 생산 기업 (중국 내수 소비 바로미터)",
            ),
            "마오타이": (
                "귀주모태주(마오타이)",
                "600519.SS",
                "필수소비재/주류(상하이CNY)",
                "🇨🇳중국 최고급 백주 '마오타이' 생산 기업 (중국 내수 소비 바로미터)",
            ),
        }
        for alias, (std_name, ticker, sec, desc) in cn_aliases.items():
            self.stock_map[alias] = {
                "std_name": std_name,
                "code": ticker,
                "market": "CN",
                "sector": sec,
                "theme_desc": desc,
            }

    def _compile_regex(self):
        keywords = sorted(self.stock_map.keys(), key=len, reverse=True)
        escaped = "|".join(map(re.escape, keywords))
        return re.compile(
            rf"(?<![A-Za-z0-9가-힣])({escaped})(?=[은는이가을를의에도과와로만\s\.,\)\]\?!]|$)"
        )

    def _extract_issue_tags(self, text: str) -> str:
        found_counts = {}
        for kw in self.issue_keywords:
            cnt = text.count(kw)
            if cnt > 0:
                found_counts[kw] = cnt
        if not found_counts:
            return "#일반시황"
        top_kws = sorted(
            found_counts.keys(), key=lambda k: found_counts[k], reverse=True
        )[:4]
        return " ".join([f"#{k}" for k in top_kws])

    def extract_stocks_with_context(self, text: str) -> list[dict]:
        unique_stocks = {}
        issue_tags = self._extract_issue_tags(text)

        for match in self.pattern.finditer(text):
            kw = match.group(1)
            info = self.stock_map[kw]
            code = info["code"]

            if code not in unique_stocks:
                start_idx = max(0, match.start() - 40)
                end_idx = min(len(text), match.end() + 50)
                snippet = text[start_idx:end_idx].strip()

                is_negative = any(
                    neg in snippet for neg in self.negative_words
                )
                tone_label = "⚠️주의/부정" if is_negative else "🟢관심/일반"

                unique_stocks[code] = {
                    "matched_word": kw,
                    "std_name": info["std_name"],
                    "code": code,
                    "market": info["market"],
                    "sector": info["sector"],
                    "theme_desc": info["theme_desc"],
                    "issue_tags": issue_tags,
                    "tone": tone_label,
                    "is_negative": is_negative,
                    "context": f"...{snippet}...",
                }
        return list(unique_stocks.values())

    def _fetch_series(self, ticker: str, min_date: datetime.date = None):
        if ticker not in self.price_cache:
            try:
                base_date = min_date if min_date else (
                    datetime.date.today() - datetime.timedelta(days=550)
                )
                start_fetch = (
                    base_date - datetime.timedelta(days=90)
                ).strftime("%Y-%m-%d")
                df = fdr.DataReader(ticker, start_fetch)
                if not df.empty:
                    df["MA20"] = df["Close"].rolling(window=20).mean()
                    df["MA60"] = df["Close"].rolling(window=60).mean()
                    delta = df["Close"].diff()
                    gain = (delta.where(delta > 0, 0)).rolling(window=14).mean()
                    loss = (
                        (-delta.where(delta < 0, 0)).rolling(window=14).mean()
                    )
                    rs = gain / loss
                    df["RSI"] = 100 - (100 / (1 + rs))
                    self.price_cache[ticker] = df
                else:
                    self.price_cache[ticker] = None
            except Exception:
                self.price_cache[ticker] = None
        return self.price_cache.get(ticker)

    def get_price_and_benchmark(
        self, ticker: str, target_date: datetime.date, market: str
    ):
        df = self._fetch_series(ticker, min_date=target_date)
        if df is None or df.empty:
            return None, None, None, None

        try:
            yesterday_str = (
                datetime.date.today() - datetime.timedelta(days=1)
            ).strftime("%Y-%m-%d")
            valid_df = df.loc[:yesterday_str]
            if valid_df.empty:
                valid_df = df

            prev_close = valid_df.iloc[-1]["Close"]

            target_str = target_date.strftime("%Y-%m-%d")
            past_df = valid_df.loc[:target_str]
            past_price = (
                past_df.iloc[-1]["Close"]
                if not past_df.empty
                else valid_df.iloc[0]["Close"]
            )

            if past_price == 0 or pd.isna(past_price) or pd.isna(prev_close):
                return None, None, None, None

            stock_yield = round(
                float(((prev_close - past_price) / past_price) * 100), 2
            )

            bm_ticker = (
                "KS11"
                if market == "KR"
                else ("US500" if market == "US" else "HSI")
            )
            bm_df = self._fetch_series(bm_ticker, min_date=target_date)
            alpha_yield = None
            if bm_df is not None and not bm_df.empty:
                bm_valid = bm_df.loc[:yesterday_str]
                if bm_valid.empty:
                    bm_valid = bm_df
                bm_past_df = bm_valid.loc[:target_str]
                bm_past = (
                    bm_past_df.iloc[-1]["Close"]
                    if not bm_past_df.empty
                    else bm_valid.iloc[0]["Close"]
                )
                bm_cur = bm_valid.iloc[-1]["Close"]
                if bm_past > 0:
                    bm_yield = ((bm_cur - bm_past) / bm_past) * 100
                    alpha_yield = round(float(stock_yield - bm_yield), 2)

            if market == "KR":
                return (
                    int(past_price),
                    int(prev_close),
                    stock_yield,
                    alpha_yield,
                )
            else:
                return (
                    round(float(past_price), 2),
                    round(float(prev_close), 2),
                    stock_yield,
                    alpha_yield,
                )
        except Exception:
            return None, None, None, None

    def analyze_trajectory(
        self, ticker: str, first_date: datetime.date, market: str
    ) -> dict:
        df = self._fetch_series(ticker, min_date=first_date)
        yesterday_str = (
            datetime.date.today() - datetime.timedelta(days=1)
        ).strftime("%Y-%m-%d")
        valid_df = (
            df.loc[:yesterday_str] if not df.loc[:yesterday_str].empty else df
        )

        first_str = first_date.strftime("%Y-%m-%d")
        period_df = valid_df.loc[first_str:]
        if period_df.empty:
            period_df = valid_df.iloc[-5:]

        cur_row = valid_df.iloc[-1]
        cur_price = cur_row["Close"]
        ma20 = cur_row.get("MA20", cur_price)
        ma60 = cur_row.get("MA60", cur_price)
        rsi = cur_row.get("RSI", 50.0)
        if pd.isna(rsi):
            rsi = 50.0

        max_price = period_df["Close"].max()
        drawdown_from_peak = (
            round(float(((cur_price - max_price) / max_price) * 100), 2)
            if max_price > 0
            else 0.00
        )

        if cur_price >= ma20 and ma20 >= ma60:
            trend_status = "📈정배열 상승추세 (20일선 위)"
        elif cur_price < ma20 and cur_price >= ma60:
            trend_status = "📐상승 중 단기눌림 (20~60일선 사이)"
        elif cur_price >= ma20 and ma20 < ma60:
            trend_status = "🌱바닥권 반등시도 (20일선 돌파)"
        else:
            trend_status = "📉하락/조정추세 (이평선 하회)"

        if rsi >= 72:
            action_rec = "⚠️단기과열 (추격매수 자제)"
        elif (
            "상승 중 단기눌림" in trend_status
            and -15.0 <= drawdown_from_peak <= -3.0
        ):
            action_rec = "⭐최우선 타점 (상승 중 건전한 눌림목)"
        elif "바닥권 반등시도" in trend_status and rsi <= 50:
            action_rec = "⭐최우선 타점 (조정 후 20일선 탈환)"
        elif rsi <= 35:
            action_rec = "💧과매도 바닥권 (분할 매수 유효)"
        elif "정배열 상승추세" in trend_status and drawdown_from_peak >= -5.0:
            action_rec = "🔥강한 추세 지속 (보유/눌림 대기)"
        else:
            action_rec = "⏳관망 (추세 회복 확인 필요)"

        peak_val = (
            int(max_price) if market == "KR" else round(float(max_price), 2)
        )
        return {
            "언급후_최고가": peak_val,
            "최고가대비_조정폭(%)": drawdown_from_peak,
            "현재_RSI": round(float(rsi), 1),
            "차트추세(이평선)": trend_status,
            "종합_진입추천": action_rec,
        }


# ==========================================
# 3. 엑셀 디자인 서식 함수
# ==========================================
def format_excel_sheet(
    writer,
    sheet_name: str,
    df: pd.DataFrame,
    market_type: str = "KR",
    freeze_cell: str = "D2",
):
    worksheet = writer.sheets[sheet_name]
    worksheet.freeze_panes = freeze_cell

    header_fill = PatternFill(
        start_color="1F4E78", end_color="1F4E78", fill_type="solid"
    )
    header_font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")

    even_row_fill = PatternFill(
        start_color="F2F6FA", end_color="F2F6FA", fill_type="solid"
    )
    odd_row_fill = PatternFill(
        start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"
    )

    default_font = Font(name="맑은 고딕", size=10, color="000000")
    bold_name_font = Font(name="맑은 고딕", size=10, bold=True, color="000000")
    red_minus_font = Font(name="맑은 고딕", size=10, bold=True, color="D91E18")

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    worksheet.row_dimensions[1].height = 26
    for col_idx in range(1, len(df.columns) + 1):
        cell = worksheet.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=False
        )
        cell.border = thin_border

    for row_idx in range(2, len(df) + 2):
        worksheet.row_dimensions[row_idx].height = 21
        row_fill = even_row_fill if row_idx % 2 == 0 else odd_row_fill

        for col_idx, col_name in enumerate(df.columns, 1):
            col_str = str(col_name)
            cell = worksheet.cell(row=row_idx, column=col_idx)
            cell.fill = row_fill
            cell.border = thin_border
            cell.font = bold_name_font if col_str == "종목명" else default_font

            val = cell.value

            if isinstance(val, (int, float)) and not pd.isna(val):
                cell.alignment = Alignment(
                    horizontal="right", vertical="center"
                )
                if (
                    "%" in col_str
                    or "수익률" in col_str
                    or "조정폭" in col_str
                    or "초과" in col_str
                ):
                    cell.number_format = "0.00"
                    if val < 0:
                        cell.font = red_minus_font
                elif "RSI" in col_str:
                    cell.number_format = "0.0"
                elif "종가" in col_str or "최고가" in col_str:
                    cell.number_format = (
                        "#,##0" if market_type == "KR" else "#,##0.00"
                    )
                elif "횟수" in col_str:
                    cell.number_format = "#,##0"
                    cell.alignment = Alignment(
                        horizontal="center", vertical="center"
                    )
            else:
                if col_str in [
                    "채널명",
                    "티커",
                    "국가",
                    "최초언급일",
                    "최근언급일",
                    "작성일",
                    "문맥분류",
                ]:
                    cell.alignment = Alignment(
                        horizontal="center", vertical="center"
                    )
                else:
                    cell.alignment = Alignment(
                        horizontal="left", vertical="center"
                    )

    for idx, col in enumerate(df.columns, 1):
        col_letter = get_column_letter(idx)
        col_name = str(col)
        max_len = sum(2 if ord(c) > 127 else 1 for c in col_name)

        for val in df[col].astype(str):
            val_len = sum(1.7 if ord(c) > 127 else 1 for c in val)
            if val_len > max_len:
                max_len = val_len

        adjusted_width = min(max(max_len + 5, 11), 62)
        worksheet.column_dimensions[col_letter].width = adjusted_width


# ==========================================
# 4. 구글 시트 연동 및 텔레그램 발송 함수
# ==========================================
def update_google_sheets(dfs_dict, sheet_title="주식_인사이트_리포트"):
    if not GOOGLE_API_AVAILABLE:
        print("▶ 구글 API 라이브러리가 없어 구글 시트 연동을 건너뜁니다.")
        return

    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if not creds_json:
        print("▶ GOOGLE_CREDENTIALS 환경 변수가 없어 구글 시트 연동을 건너뜁니다.")
        return

    try:
        print(f"▶ 온라인 구글 시트('{sheet_title}') 업데이트를 시작합니다...")
        creds_dict = json.loads(creds_json)
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = Credentials.from_service_account_info(
            creds_dict, scopes=scope
        )
        client = gspread.authorize(creds)

        spreadsheet = client.open(sheet_title)

        for sheet_name, df in dfs_dict.items():
            if df.empty:
                continue
            try:
                worksheet = spreadsheet.worksheet(sheet_name)
            except gspread.exceptions.WorksheetNotFound:
                worksheet = spreadsheet.add_worksheet(
                    title=sheet_name, rows=100, cols=20
                )

            worksheet.clear()
            # 날짜 및 숫자 데이터를 문자열로 변환하여 구글 시트에 깔끔하게 기록
            df_str = df.astype(str)
            data_to_upload = [df_str.columns.tolist()] + df_str.values.tolist()
            worksheet.update(data_to_upload)
            print(f"   - '{sheet_name}' 시트 업로드 완료!")
    except Exception as e:
        print(f"❌ 구글 시트 업데이트 중 오류 발생: {e}")


def send_telegram_alert(output_file: str, summary_text: str):
    token = os.environ.get("TELEGRAM_TOKEN")
    chat_id = os.environ.get("TELEGRAM_CHAT_ID")
    if not token or not chat_id:
        print("▶ 텔레그램 토큰 또는 채팅방 ID가 없어 알림 전송을 건너뜁니다.")
        return

    try:
        print("▶ 텔레그램으로 요약 메시지와 엑셀 파일을 전송합니다...")
        # 1) 텍스트 메시지 전송
        msg_url = f"https://api.telegram.org/bot{token}/sendMessage"
        requests.post(
            msg_url,
            data={
                "chat_id": chat_id,
                "text": summary_text,
                "parse_mode": "Markdown",
            },
        )

        # 2) 엑셀 파일 첨부 전송
        doc_url = f"https://api.telegram.org/bot{token}/sendDocument"
        with open(output_file, "rb") as f:
            requests.post(
                doc_url,
                data={"chat_id": chat_id, "caption": "📊 오늘의 주식 타점 분석 리포트"},
                files={"document": f},
            )
        print("▶ 텔레그램 전송 완료!")
    except Exception as e:
        print(f"❌ 텔레그램 전송 중 오류 발생: {e}")


# ==========================================
# 5. 메인 실행 파이프라인
# ==========================================
DB_FILE = "stock_mention_history_db_v4.csv"


def run_pipeline():
    INITIAL_SEARCH_DAYS = 365  # 운영 시 365로 변경 가능

    if os.path.exists(DB_FILE):
        df_history = pd.read_csv(DB_FILE, dtype={"content_id": str, "티커": str})
        df_history["작성일"] = pd.to_datetime(df_history["작성일"]).dt.date
        known_ids = set(df_history["content_id"].unique())
    else:
        df_history = pd.DataFrame()
        known_ids = set()

    collectors = [
        NaverBlogCollector(blog_id="ranto28", blog_name="메르"),
    ]

    new_contents = []
    for collector in collectors:
        new_contents.extend(
            collector.collect(
                days_limit=INITIAL_SEARCH_DAYS, known_ids=known_ids
            )
        )

    engine = StockEngine()

    if new_contents:
        new_records = []
        for item in new_contents:
            extracted = engine.extract_stocks_with_context(item.text)
            for stock in extracted:
                new_records.append({
                    "content_id": str(item.content_id),
                    "채널명": item.channel_name,
                    "작성일": item.published_date,
                    "국가": stock["market"],
                    "종목명": stock["std_name"],
                    "티커": str(stock["code"]),
                    "업종_섹터": stock["sector"],
                    "핵심테마_하는일": stock["theme_desc"],
                    "이슈_키워드": stock["issue_tags"],
                    "문맥분류": stock["tone"],
                    "본문_언급_문맥": stock["context"],
                    "콘텐츠_제목": item.title,
                    "링크": item.url,
                    "is_negative": stock["is_negative"],
                })
        if new_records:
            df_new = pd.DataFrame(new_records)
            df_history = pd.concat([df_history, df_new], ignore_index=True)
            df_history.drop_duplicates(
                subset=["채널명", "content_id", "티커"], inplace=True
            )
            df_history.to_csv(DB_FILE, index=False, encoding="utf-8-sig")
            print(f"▶ 신규 종목 언급 {len(new_records)}건이 장부에 추가되었습니다!")

    if df_history.empty:
        print("누적된 종목 데이터가 없습니다.")
        return

    updated_rows = []
    for _, row in df_history.iterrows():
        past_p, prev_close_p, rate, alpha = engine.get_price_and_benchmark(
            str(row["티커"]), row["작성일"], row["국가"]
        )
        if past_p is not None:
            row_dict = row.to_dict()
            latest_info = engine.stock_map.get(row["종목명"], {})
            if latest_info:
                row_dict["업종_섹터"] = latest_info.get(
                    "sector", row.get("업종_섹터", "")
                )
                row_dict["핵심테마_하는일"] = latest_info.get(
                    "theme_desc", row.get("핵심테마_하는일", "")
                )

            row_dict["언급일_종가"] = past_p
            row_dict["전일_종가"] = prev_close_p
            row_dict["전일종가_수익률(%)"] = rate
            row_dict["시장초과수익률(%p)"] = alpha
            updated_rows.append(row_dict)

    df_all = pd.DataFrame(updated_rows)
    df_pos = df_all[~df_all["is_negative"]].sort_values(
        by="작성일", ascending=True
    )

    summary_rows = []
    for (channel, market, name, ticker), group in df_pos.groupby(
        ["채널명", "국가", "종목명", "티커"]
    ):
        first_row = group.iloc[0]
        last_row = group.iloc[-1]
        traj = engine.analyze_trajectory(ticker, first_row["작성일"], market)

        summary_rows.append({
            "채널명": channel,
            "종목명": name,
            "티커": ticker,
            "국가": market,
            "핵심테마_하는일": last_row.get("핵심테마_하는일", ""),
            "최근_이슈키워드": last_row.get("이슈_키워드", ""),
            "관심언급횟수": len(group),
            "종합_진입추천": traj["종합_진입추천"],
            "차트추세(이평선)": traj["차트추세(이평선)"],
            "최초언급일": first_row["작성일"],
            "최초언급_종가": first_row["언급일_종가"],
            "언급후_최고가": traj["언급후_최고가"],
            "전일_종가": last_row["전일_종가"],
            "최초대비_수익률(%)": first_row["전일종가_수익률(%)"],
            "최고가대비_조정폭(%)": traj["최고가대비_조정폭(%)"],
            "현재_RSI": traj["현재_RSI"],
            "최근언급일": last_row["작성일"],
            "최근언급_문맥": last_row["본문_언급_문맥"],
            "최근글_링크": last_row["링크"],
        })

    df_summary = pd.DataFrame(summary_rows)
    if not df_summary.empty:
        df_summary = df_summary.sort_values(
            by=["최근언급일", "관심언급횟수"], ascending=[False, False]
        )

    df_kr = (
        df_summary[df_summary["국가"] == "KR"].drop(columns=["국가"])
        if not df_summary.empty
        else pd.DataFrame()
    )
    df_us = (
        df_summary[df_summary["국가"] == "US"].drop(columns=["국가"])
        if not df_summary.empty
        else pd.DataFrame()
    )
    df_cn = (
        df_summary[df_summary["국가"] == "CN"].drop(columns=["국가"])
        if not df_summary.empty
        else pd.DataFrame()
    )

    df_detail = df_all.drop(columns=["content_id", "is_negative"]).sort_values(
        by="작성일", ascending=False
    )
    detail_cols = [
        "채널명",
        "종목명",
        "티커",
        "국가",
        "작성일",
        "업종_섹터",
        "핵심테마_하는일",
        "이슈_키워드",
        "문맥분류",
        "언급일_종가",
        "전일_종가",
        "전일종가_수익률(%)",
        "시장초과수익률(%p)",
        "본문_언급_문맥",
        "콘텐츠_제목",
        "링크",
    ]
    df_detail = df_detail[detail_cols]

    # 엑셀 파일 저장
    output_file = f"주식인사이트리포트_{datetime.date.today().strftime('%Y%m%d')}.xlsx"
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        if not df_kr.empty:
            df_kr.to_excel(
                writer, sheet_name="1_한국주식_타점(원)", index=False
            )
            format_excel_sheet(
                writer,
                "1_한국주식_타점(원)",
                df_kr,
                market_type="KR",
                freeze_cell="D2",
            )
        if not df_us.empty:
            df_us.to_excel(
                writer, sheet_name="2_미국주식_타점(달러)", index=False
            )
            format_excel_sheet(
                writer,
                "2_미국주식_타점(달러)",
                df_us,
                market_type="US",
                freeze_cell="D2",
            )
        if not df_cn.empty:
            df_cn.to_excel(
                writer, sheet_name="3_중국홍콩_타점", index=False
            )
            format_excel_sheet(
                writer,
                "3_중국홍콩_타점",
                df_cn,
                market_type="CN",
                freeze_cell="D2",
            )
        df_detail.to_excel(
            writer, sheet_name="4_전체_누적언급내역", index=False
        )
        format_excel_sheet(
            writer,
            "4_전체_누적언급내역",
            df_detail,
            market_type="ALL",
            freeze_cell="D2",
        )

    # 온라인 구글 시트 업데이트 실행
    dfs_for_google = {}
    if not df_kr.empty:
        dfs_for_google["1_한국주식_타점(원)"] = df_kr
    if not df_us.empty:
        dfs_for_google["2_미국주식_타점(달러)"] = df_us
    if not df_cn.empty:
        dfs_for_google["3_중국홍콩_타점"] = df_cn
    dfs_for_google["4_전체_누적언급내역"] = df_detail

    update_google_sheets(dfs_for_google, sheet_title="Stock Insight Report")

    # 텔레그램 알림 메시지 생성
    top_picks = (
        df_summary[df_summary["종합_진입추천"].str.contains("⭐최우선 타점")]
        .head(3)
        .to_dict("records")
    )
    pick_str = (
        "\n".join([
            f"• **{p['종목명']}** ({p['티커']}): {p['종합_진입추천']}"
            for p in top_picks
        ])
        if top_picks
        else "• 특이 포착 종목 없음"
    )

    telegram_msg = (
        f"📊 *[메르 블로그 주식 타점 리포트]*\n"
        f"📅 기준일: {datetime.date.today()}\n\n"
        f"🎯 *주요 타점 포착 종목 추천*\n{pick_str}\n\n"
        f"🔗 구글 시트(`Stock Insight Report`) 및 첨부된 엑셀 파일에서 상세 지표(RSI, 이평선 추세, 최고가 대비 조정폭)를 확인하세요!"
    )
    send_telegram_alert(output_file, telegram_msg)


if __name__ == "__main__":
    run_pipeline()