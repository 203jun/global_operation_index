#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LMS Learning Excel 파일 전처리 스크립트
LMS Learning Excel 파일의 컬럼 리스트를 추출하는 기능을 제공합니다.
"""

import pandas as pd
import os
import sys
from logger_config import get_default_logger

# 로거 설정
logger = get_default_logger(__name__)

def load_excel_file(file_path):
    """
    Excel 파일을 불러오는 공통 함수

    Args:
        file_path (str): Excel 파일 경로

    Returns:
        pandas.ExcelFile: Excel 파일 객체
    """
    logger.info("1단계: LMS Learning Excel 파일을 불러옵니다...")
    xl_file = pd.ExcelFile(file_path)
    logger.info("✓ Excel 파일 불러오기 완료")
    return xl_file

def get_lms_columns(file_path):
    """
    Excel 파일에서 모든 컬럼명을 반환하는 함수

    Args:
        file_path (str): Excel 파일 경로

    Returns:
        list: 컬럼명 리스트
    """
    try:
        # 1단계: Excel 파일 불러오기
        xl_file = load_excel_file(file_path)

        # 2단계: 컬럼명 확인
        logger.info("2단계: 컬럼명을 확인합니다...")
        df = pd.read_excel(file_path, nrows=0)
        columns = list(df.columns)

        logger.info(f"✓ 컬럼명 확인 완료: 총 {len(columns)}개 컬럼")
        logger.info(f"컬럼 리스트: {columns}")

        return columns

    except Exception as e:
        logger.error(f"✗ 오류 발생: {e}")
        return None

def group_category(file_path):
    """
    Category를 그룹핑하는 함수

    Args:
        file_path (str): Excel 파일 경로

    Returns:
        pandas.DataFrame: category_big 컬럼이 추가된 데이터프레임
    """
    try:
        # 3단계: Category 그룹핑
        logger.info("3단계: Category를 그룹핑합니다...")

        # 데이터 읽기
        df = pd.read_excel(file_path)
        logger.info(f"✓ 데이터 읽기 완료: {df.shape[0]}행, {df.shape[1]}열")

        # Category 컬럼 찾기
        category_col = None
        for col in df.columns:
            if 'category' in col.lower():
                category_col = col
                break

        if category_col is None:
            logger.error("✗ Category 컬럼을 찾을 수 없습니다.")
            return None

        logger.info(f"✓ Category 컬럼: {category_col}")

        # 3.1단계: 실제 Category 값들 확인
        logger.info("3.1단계: 실제 Category 값들을 확인합니다...")
        unique_categories = df[category_col].dropna().unique()
        logger.info(f"✓ 실제 Category 값들 ({len(unique_categories)}개): {list(unique_categories)}")

        # 3.2단계: Category 매핑
        logger.info("3.2단계: Category를 매핑합니다...")
        logger.info("  - 'category_1' (상위 카테고리)와 'category_2' (하위 카테고리) 컬럼 생성")
        logger.info("  - 원본 데이터: 'Category' 컬럼 값 사용")
        logger.info("  - 매핑 규칙:")
        logger.info("    '경력사원, 신규입사자, 신입사원' → ('신입온보딩', '신입온보딩')")
        logger.info("    '고객 가치, 고객마인드' → ('직무', '고객가치')")
        logger.info("    '고객 가치, 고객중심 일하는 방식' → ('직무', '고객가치')")
        logger.info("    '구매관리' → ('직무', '구매')")
        logger.info("    '노경' → ('직무', 'HR')")
        logger.info("    '독서통신' → ('직무공통', '직무공통')")
        logger.info("    '리더십', '리더십 공통, 직무역량', ... → ('리더십', '일반')")
        logger.info("    '리더십, 직책 리더십, 파트장/팀장', ... → ('리더십', '직책')")
        logger.info("    '마케팅, 영업, 직무역량' → ('직무', '마케팅/영업')")
        logger.info("    '보안관리', '업무시스템', '직무공통', ... → ('직무공통', '직무공통')")
        logger.info("    '생산관리', '생산기술', 'Production R&D' → ('직무', '생산')")
        logger.info("    '소재 R&D', 'R&D 공통', 'Software R&D', ... → ('직무', 'R&D')")
        logger.info("    '신규입사자', '신규입사자, 신입사원', ... → ('신입온보딩', '신입온보딩')")
        logger.info("    '영업', '영업, 직무역량', ... → ('직무', '마케팅/영업')")
        logger.info("    '자재', '제조' → ('직무', '자재/제조')")
        logger.info("    '재경' → ('직무', '재경')")
        logger.info("    '전략기획' → ('직무공통', '직무공통')")
        logger.info("    '품질', '품질관리', ... → ('직무', '품질')")
        logger.info("    'HR', 'L&D', '조직문화', ... → ('직무', 'HR')")
        logger.info("    '직무역량, 직책 리더십, 휴넷', '파트장/팀장', ... → ('리더십', '직책')")
        logger.info("    '핵심인재' → ('리더십', '핵심인재')")
        logger.info("    '환경안전', 'LG 필수 교육' → ('전사필수', '전사필수')")
        logger.info("    'AI/빅데이터', 'DX', ... → ('직무', 'AI/DX')")
        logger.info("    'B2B', 'B2B영업', ... → ('직무', 'B2B')")
        logger.info("    'LG 경영방침', 'LG사례', ... → ('직무공통', 'LG')")
        logger.info("    'SCM' → ('직무', 'SCM')")
        logger.info("    기타 매핑되지 않은 값 → ('UNMAPPED', 원본값)")

        # Category 매핑 함수 (제공된 매핑 규칙 기반)
        def map_categories(category):
            if pd.isna(category):
                return '기타', '기타'

            category_str = str(category)

            # 제공된 매핑 규칙 적용
            if category_str in ['경력사원, 신규입사자, 신입사원']:
                return '신입온보딩', '신입온보딩'
            elif category_str in ['고객 가치, 고객마인드', '고객 가치, 고객중심 일하는 방식']:
                return '직무', '고객가치'
            elif category_str in ['구매관리']:
                return '직무', '구매'
            elif category_str in ['노경']:
                return '직무', 'HR'
            elif category_str in ['독서통신']:
                return '직무공통', '직무공통'
            elif category_str in ['리더십', '리더십 공통, 직무역량', '리더십 공통, 직무역량, 휴넷', '리더십 공통, 휴넷', '리더십 기타', '리더십, 리더십 공통', '리더십, 직무역량']:
                return '리더십', '일반'
            elif category_str in ['리더십, 직책 리더십, 파트장/팀장', '리더십, 파트장/팀장']:
                return '리더십', '직책'
            elif category_str in ['마케팅, 영업, 직무역량']:
                return '직무', '마케팅/영업'
            elif category_str in ['보안관리', '비즈니스 기본스킬, 직무역량', '사별특화영역', '산업 연수', '성과 모니터링', '업무시스템', '직무공통', '직무역량', 'IT 기본, Security', 'RPA', 'Security']:
                return '직무공통', '직무공통'
            elif category_str in ['생산관리', '생산기술', 'Production R&D']:
                return '직무', '생산'
            elif category_str in ['소재 R&D', '시스템 SW', 'Hardware R&D, 직무역량', 'Hardware R&D, Software R&D, 기구 R&D, 직무역량, 품질', 'R&D 공통', 'R&D 공통, 영업, 직무역량', 'R&D 공통, 직무역량', 'R&D 공통, 직무역량, 품질', 'R&D 공통, Software R&D, 직무역량', 'R&D기획/관리', 'Software R&D']:
                return '직무', 'R&D'
            elif category_str in ['신규입사자', '신규입사자, 신입사원', '신규입사자, 영업, 직무역량']:
                return '신입온보딩', '신입온보딩'
            elif category_str in ['영업', '영업, 직무역량', '영업, 직무역량, 품질']:
                return '직무', '마케팅/영업'
            elif category_str in ['자재', '제조']:
                return '직무', '자재/제조'
            elif category_str in ['재경']:
                return '직무', '재경'
            elif category_str in ['전략기획']:
                return '직무공통', '직무공통'
            elif category_str in ['제품설계, 직무역량, 품질', '직무역량, 품질', '품질', '품질관리']:
                return '직무', '품질'
            elif category_str in ['조직문화', 'HR', 'HR, 조직문화', 'HR, L&D', 'HRM', 'L&D', 'L&D, 노경, 조직문화', 'L&D, 신규입사자', 'L&D, 신입사원', 'L&D, 직무역량']:
                return '직무', 'HR'
            elif category_str in ['직무역량, 직책 리더십, 휴넷', '직무역량, 파트장/팀장', '파트장/팀장']:
                return '리더십', '직책'
            elif category_str in ['핵심인재']:
                return '리더십', '핵심인재'
            elif category_str in ['환경안전', 'LG 필수 교육']:
                return '전사필수', '전사필수'
            elif category_str in ['AI/빅데이터', 'AI/빅데이터, DX', 'AI/빅데이터, DX, DX Technology, DX 사례연구, Digital Literacy, LG사례, 데이터분석, 빅데이터, 인공지능, 품질, 품질관리, 프로그래밍', 'AI/빅데이터, DX, DX Technology, DX 사례연구, Digital Literacy, LG사례, 데이터분석, 빅데이터, 인공지능, 프로그래밍', 'AI/빅데이터, DX, DX Technology, DX 사례연구, Digital Literacy, LG사례, 분석, 빅데이터, 상품기획, 인공지능', 'AI/빅데이터, DX, DX Technology, DX 사례연구, Digital Literacy, LG사례, 빅데이터, 인공지능, 품질', 'DX, 인공지능', 'DX, DX 사례연구, LG사례, 글로벌사례, 빅데이터, 인공지능', 'DX, DX Technology, DX 사례연구, Digital Literacy, LG사례, 데이터분석, 빅데이터, 인공지능, 통계, 프로그래밍']:
                return '직무', 'AI/DX'
            elif category_str in ['B2B', 'B2B, 영업, 직무역량', 'B2B, B2B영업', 'B2B영업']:
                return '직무', 'B2B'
            elif category_str in ['LG 경영방침, LG 리더의 사업철학', 'LG 리더의 사업철학', 'LG사례']:
                return '직무공통', 'LG'
            elif category_str in ['SCM']:
                return '직무', 'SCM'
            else:
                # 매핑되지 않은 항목
                return 'UNMAPPED', category_str

        # category_1, category_2 컬럼 생성
        logger.info("  - 매핑 작업 수행 중...")
        mapped_categories = df[category_col].apply(map_categories)
        df['category_1'] = [cat[0] for cat in mapped_categories]
        df['category_2'] = [cat[1] for cat in mapped_categories]
        logger.info(f"    ✓ 'category_1' 컬럼 생성 완료: {df['category_1'].nunique()}개 고유값")
        logger.info(f"    ✓ 'category_2' 컬럼 생성 완료: {df['category_2'].nunique()}개 고유값")

        # 매핑 결과 요약
        total_count = len(df)
        category_1_counts = df['category_1'].value_counts()
        logger.info("  - category_1 (상위 카테고리) 생성 결과:")
        for category, count in category_1_counts.items():
            percentage = (count / total_count * 100) if total_count > 0 else 0
            logger.info(f"    {category}: {count}건 ({percentage:.1f}%)")
        logger.info(f"    전체: {total_count}건")

        category_2_counts = df['category_2'].value_counts()
        logger.info("  - category_2 (하위 카테고리) 생성 결과:")
        for category, count in category_2_counts.items():
            percentage = (count / total_count * 100) if total_count > 0 else 0
            logger.info(f"    {category}: {count}건 ({percentage:.1f}%)")
        logger.info(f"    전체: {total_count}건")

        logger.info("✓ 3.2단계 완료: Category 매핑 완료")

        # 3.3단계: 매핑되지 않은 항목 확인
        logger.info("3.3단계: 매핑되지 않은 항목을 확인합니다...")
        unmapped_items = df[df['category_1'] == 'UNMAPPED']['category_2'].unique()
        if len(unmapped_items) > 0:
            logger.warning("매핑되지 않은 Category 항목들:")
            for item in unmapped_items:
                logger.warning(f"  - {item}")
        else:
            logger.info("✓ 매핑되지 않은 항목이 없습니다.")

        logger.info("✓ Category 그룹핑 완료")

        return df

    except Exception as e:
        logger.error(f"✗ 오류 발생: {e}")
        return None
