#!/usr/bin/env python3
"""
HR Index Excel 파일 전처리 스크립트
HR Index Excel 파일에서 detail sheet의 컬럼 리스트를 추출하는 기능을 제공합니다.
"""

import pandas as pd
import os
import sys
from logger_config import get_default_logger

# 로거 설정
logger = get_default_logger(__name__)

def get_sheet_columns(file_path, sheet_name_keyword):
    """
    Excel 파일에서 특정 키워드가 포함된 시트의 컬럼 리스트를 추출하는 함수

    Args:
        file_path (str): Excel 파일 경로
        sheet_name_keyword (str): 시트 이름에 포함될 키워드

    Returns:
        tuple: (시트명, 컬럼 리스트) 또는 (None, None)
    """
    try:
        # Excel 파일 읽기
        xl_file = pd.ExcelFile(file_path)
        logger.info(f"Available sheets: {xl_file.sheet_names}")

        # 키워드가 포함된 시트 찾기 (대소문자 구분 없이)
        target_sheet = None
        for sheet_name in xl_file.sheet_names:
            if sheet_name_keyword.lower() in sheet_name.lower():
                target_sheet = sheet_name
                break

        if target_sheet is None:
            logger.warning(f"{sheet_name_keyword} 시트를 찾을 수 없습니다.")
            logger.info("사용 가능한 시트:")
            for i, sheet in enumerate(xl_file.sheet_names):
                logger.info(f"  {i+1}. {sheet}")
            return None, None

        logger.info(f"Found {sheet_name_keyword} sheet: {target_sheet}")

        # 시트 데이터 읽기 (첫 번째 행을 컬럼으로 사용)
        df = pd.read_excel(file_path, sheet_name=target_sheet, header=0)
        columns = list(df.columns)

        logger.info(f"{target_sheet} sheet columns ({len(columns)}): {columns}")

        # 컬럼이 0개인 경우, 실제 데이터가 있는지 확인
        if len(columns) == 0:
            logger.warning(f"{target_sheet} 시트에 컬럼이 없습니다. 실제 데이터 확인 중...")
            # 첫 5행 읽어서 확인
            df_sample = pd.read_excel(file_path, sheet_name=target_sheet, nrows=5)
            logger.info(f"Sample data shape: {df_sample.shape}")
            if not df_sample.empty:
                logger.info(f"Sample data:\n{df_sample}")
                # 첫 번째 행을 컬럼으로 사용할 수 있는지 확인
                if len(df_sample.columns) > 0:
                    columns = list(df_sample.columns)
                    logger.info(f"Using first row as columns: {columns}")

        return target_sheet, columns

    except Exception as e:
        logger.error(f"Error processing file: {e}")
        return None, None


def load_excel_file(file_path):
    """
    Excel 파일을 불러오는 공통 함수

    Args:
        file_path (str): Excel 파일 경로

    Returns:
        pandas.ExcelFile: Excel 파일 객체
    """
    logger.info("1단계: hr_index.xlsx 파일을 불러옵니다...")
    xl_file = pd.ExcelFile(file_path)
    logger.info("✓ Excel 파일 불러오기 완료")
    return xl_file

def find_detail_sheet(xl_file):
    """
    DETAIL 시트를 찾는 공통 함수

    Args:
        xl_file: Excel 파일 객체

    Returns:
        str: DETAIL 시트명
    """
    logger.info("2단계: DETAIL 시트를 찾습니다...")
    logger.info(f"  - 사용 가능한 시트: {xl_file.sheet_names}")

    detail_sheet = None
    for sheet_name in xl_file.sheet_names:
        if sheet_name.lower() == 'detail':
            detail_sheet = sheet_name
            logger.info(f"  ✓ DETAIL 시트 발견: '{sheet_name}'")
            break

    if detail_sheet is None:
        logger.error(f"✗ DETAIL 시트를 찾을 수 없습니다. 사용 가능한 시트: {xl_file.sheet_names}")
        return None

    return detail_sheet

def get_detail_columns(file_path, detail_sheet):
    """
    Excel 파일에서 detail sheet의 모든 컬럼명을 반환하는 함수

    Args:
        file_path (str): Excel 파일 경로
        detail_sheet (str): DETAIL 시트명

    Returns:
        list: detail sheet의 컬럼명 리스트
    """
    try:
        # 3단계: 컬럼명 확인
        logger.info("3단계: DETAIL 시트의 컬럼명을 확인합니다...")
        df = pd.read_excel(file_path, sheet_name=detail_sheet, nrows=0)
        columns = list(df.columns)

        logger.info(f"✓ 컬럼명 확인 완료: 총 {len(columns)}개 컬럼")
        logger.info(f"컬럼 리스트: {columns}")

        return columns

    except Exception as e:
        logger.error(f"✗ 오류 발생: {e}")
        return None

def create_final_company_name(file_path, detail_sheet):
    """
    최종 법인명 컬럼을 생성하는 함수

    Args:
        file_path (str): Excel 파일 경로
        detail_sheet (str): DETAIL 시트명

    Returns:
        pandas.DataFrame: 최종 법인명이 추가된 데이터프레임
    """
    try:
        # 4단계: 최종 법인명 생성
        logger.info("4단계: 최종 법인명을 생성합니다...")

        # 4.1단계: 데이터 읽기
        logger.info("4.1단계: DETAIL 시트 데이터를 읽습니다...")
        df = pd.read_excel(file_path, sheet_name=detail_sheet)
        logger.info(f"✓ 데이터 읽기 완료: {df.shape[0]}행, {df.shape[1]}열")

        # 4.2단계: 컬럼 찾기
        logger.info("4.2단계: 필요한 컬럼을 찾습니다...")
        logger.info(f"  - 전체 컬럼 수: {len(df.columns)}")
        logger.info(f"  - 전체 컬럼 목록: {list(df.columns)[:10]}..." if len(df.columns) > 10 else f"  - 전체 컬럼 목록: {list(df.columns)}")

        # Integrated Sub. Name(MP) 컬럼 찾기 (정확한 컬럼명으로 검색)
        logger.info("  - 'Integrated Sub. Name(MP)' 컬럼 검색 중...")
        integrated_col = 'Integrated Sub. Name(MP)' if 'Integrated Sub. Name(MP)' in df.columns else None

        if integrated_col:
            logger.info(f"    ✓ 컬럼 발견: '{integrated_col}'")
        else:
            logger.warning("    ✗ 'Integrated Sub. Name(MP)' 컬럼을 찾지 못했습니다.")

        # Sub. Name(MP) 컬럼 찾기 (정확한 컬럼명으로 검색)
        logger.info("  - 'Sub. Name(MP)' 컬럼 검색 중...")
        sub_name_col = 'Sub. Name(MP)' if 'Sub. Name(MP)' in df.columns else None

        if sub_name_col:
            logger.info(f"    ✓ 컬럼 발견: '{sub_name_col}'")
        else:
            logger.warning("    ✗ 'Sub. Name(MP)' 컬럼을 찾지 못했습니다.")

        if integrated_col is None or sub_name_col is None:
            logger.error("✗ 필요한 컬럼을 찾을 수 없습니다.")
            logger.error(f"  - Integrated Sub. Name(MP): {'✓ 찾음' if integrated_col else '✗ 못찾음'}")
            logger.error(f"  - Sub. Name(MP): {'✓ 찾음' if sub_name_col else '✗ 못찾음'}")
            return None

        logger.info("✓ 4.2단계 완료: 필요한 컬럼을 모두 찾았습니다.")

        # 4.3단계: 최종 법인명 생성
        logger.info("4.3단계: 법인명 로직을 적용합니다...")
        logger.info("  법인명 결정 로직:")
        logger.info("    1) 'Integrated Sub. Name(MP)'가 null이 아니면 → 그 값을 사용")
        logger.info("    2) 'Integrated Sub. Name(MP)'가 null이고, 'Sub. Name(MP)'가 null이 아니면 → 'Sub. Name(MP)' 값을 사용")
        logger.info("    3) 둘 다 null이면 → 'ETC'를 사용")

        def get_final_company_name(row):
            # Integrated Sub. Name(MP)가 null이 아니면 그대로 사용
            if pd.notna(row[integrated_col]) and str(row[integrated_col]).strip() != '':
                return row[integrated_col]
            # Sub. Name(MP)가 null이 아니면 사용
            elif pd.notna(row[sub_name_col]) and str(row[sub_name_col]).strip() != '':
                return row[sub_name_col]
            # 둘 다 null이면 ETC
            else:
                return 'ETC'

        logger.info("  - 법인명 로직 적용 중...")
        df['Final Sub.'] = df.apply(get_final_company_name, axis=1)

        # 로직 적용 결과 통계
        integrated_used = len(df[pd.notna(df[integrated_col]) & (df[integrated_col].astype(str).str.strip() != '')])
        sub_name_used = len(df[(df[integrated_col].isna() | (df[integrated_col].astype(str).str.strip() == '')) &
                                pd.notna(df[sub_name_col]) & (df[sub_name_col].astype(str).str.strip() != '')])
        etc_used = len(df[df['Final Sub.'] == 'ETC'])

        logger.info(f"  - 로직 적용 결과:")
        logger.info(f"    → Integrated Sub. Name(MP) 사용: {integrated_used}건")
        logger.info(f"    → Sub. Name(MP) 사용: {sub_name_used}건")
        logger.info(f"    → ETC 사용: {etc_used}건")
        logger.info(f"    → 전체: {len(df)}건")

        # 4.4단계: Finland Lab → LGEFL 변경
        logger.info("4.4단계: Finland Lab을 LGEFL로 변경합니다...")
        finland_count = len(df[df[integrated_col] == 'Finland Lab'])
        df.loc[df[integrated_col] == 'Finland Lab', 'Final Sub.'] = 'LGEFL'
        logger.info(f"✓ Finland Lab → LGEFL 변경 완료: {finland_count}개")

        # 4.5단계: Monterrey Factory → LGEMN 변경
        logger.info("4.5단계: Monterrey Factory를 LGEMN으로 변경합니다...")
        monterrey_condition = (df[integrated_col].isna() | (df[integrated_col].astype(str).str.strip() == '')) & (df[sub_name_col] == 'Monterrey Factory')
        monterrey_count = len(df[monterrey_condition])
        df.loc[monterrey_condition, 'Final Sub.'] = 'LGEMN'
        logger.info(f"✓ Monterrey Factory → LGEMN 변경 완료: {monterrey_count}개")

        # 4.6단계: ~~Branch로 끝나는 경우 Final Sub.에 그대로 사용
        logger.info("4.6단계: ~~Branch로 끝나는 경우 Final Sub.에 그대로 적용합니다...")
        branch_condition = df[sub_name_col].astype(str).str.endswith('Branch')
        branch_count = len(df[branch_condition])
        if branch_count > 0:
            df.loc[branch_condition, 'Final Sub.'] = df.loc[branch_condition, sub_name_col]
            logger.info(f"✓ ~~Branch 이름을 Final Sub.에 적용 완료: {branch_count}개")

            # 적용된 Branch 이름들 출력
            applied_branches = df.loc[branch_condition, sub_name_col].unique()
            logger.info("적용된 Branch 이름들:")
            for branch in applied_branches:
                count = len(df[df[sub_name_col] == branch])
                logger.info(f"  {branch}: {count}개")
        else:
            logger.info("✓ ~~Branch로 끝나는 데이터가 없습니다.")

        # 4.7단계: Final Sub. 대문자 변환
        logger.info("4.7단계: Final Sub.를 대문자로 변환합니다...")
        logger.info("  - 목적: 다른 데이터(LMS, HONG)와 조인 시 키 매칭을 정확하게 하기 위함")
        logger.info("  - 이유: 대소문자 차이로 인한 조인 실패 방지")
        logger.info("    예: 'lgejp'와 'LGEJP'가 다른 법인으로 인식되는 문제 방지")
        logger.info("  - 방법: 모든 Final Sub. 값을 대문자로 통일")

        # 변환 전 샘플
        before_sample = df['Final Sub.'].head(5).tolist()
        logger.info(f"  - 변환 전 샘플: {before_sample}")

        # 대문자 변환
        df['Final Sub.'] = df['Final Sub.'].str.upper()

        # 변환 후 샘플
        after_sample = df['Final Sub.'].head(5).tolist()
        logger.info(f"  - 변환 후 샘플: {after_sample}")
        logger.info("✓ 4.7단계 완료: Final Sub. 대문자 변환 완료")

        # Final Sub. 컬럼 생성 요약
        logger.info("✓ 4단계 완료: 'Final Sub.' 컬럼 생성 완료")
        logger.info("  - 컬럼 생성 과정 요약:")
        logger.info("    1) 기본 로직으로 Integrated Sub. Name(MP) 또는 Sub. Name(MP) 사용")
        logger.info("    2) Finland Lab → LGEFL 특수 처리")
        logger.info("    3) Monterrey Factory → LGEMN 특수 처리")
        logger.info("    4) ~~Branch로 끝나는 경우 Sub. Name(MP) 그대로 사용")
        logger.info("    5) 모든 법인명을 대문자로 변환 (조인 키 통일)")

        # 결과 요약
        final_counts = df['Final Sub.'].value_counts()
        logger.info(f"  - 최종 고유 법인명: {len(final_counts)}개")
        logger.info("  - 법인별 데이터 건수 (상위 10개):")
        for company, count in final_counts.head(10).items():
            logger.info(f"    {company}: {count}건")

        return df

    except Exception as e:
        logger.error(f"✗ 오류 발생: {e}")
        return None

def update_region_mp_complete(df_input, file_path, detail_sheet):
    """
    Region(MP) 컬럼을 수정하는 완전한 함수 (5단계)
    """
    try:
        # 5단계: Region(MP) 수정
        logger.info("5단계: Region(MP)을 수정합니다...")

        # 데이터 읽기 (df_input을 사용하도록 변경)
        df = df_input.copy()  # 원본 데이터프레임 변경 방지

        # Region(MP) 컬럼 찾기
        region_col = None
        for col in df.columns:
            if 'region' in col.lower() and 'mp' in col.lower():
                region_col = col
                break

        if region_col is None:
            logger.error("✗ Region(MP) 컬럼을 찾을 수 없습니다.")
            return None

        logger.info(f"✓ Region(MP) 컬럼: {region_col}")

        # Integrated Sub. Name(MP) 컬럼 찾기
        integrated_col = None
        for col in df.columns:
            if 'integrated' in col.lower() and 'sub' in col.lower() and 'name' in col.lower() and 'mp' in col.lower():
                integrated_col = col
                break

        if integrated_col is None:
            logger.error("✗ Integrated Sub. Name(MP) 컬럼을 찾을 수 없습니다.")
            return None

        # Final Region 컬럼 생성
        logger.info("  - 'Final Region' 컬럼 생성 중...")
        logger.info(f"    기본값: 'Region(MP)' 컬럼 ({region_col}) 값을 복사")
        df['Final Region'] = df[region_col].copy()

        # 생성된 Final Region의 고유값 확인
        initial_regions = df['Final Region'].unique()
        logger.info(f"    생성된 Final Region 고유값 ({len(initial_regions)}개): {list(initial_regions)}")

        # LGECH + Asia Region → China 변경
        logger.info("  - 특수 케이스 처리: LGECH + Asia Region → China 변경")
        logger.info("    조건: Integrated Sub. Name(MP) = 'LGECH' AND Region(MP) = 'Asia Region'")
        lgech_condition = (df[integrated_col] == 'LGECH') & (df[region_col] == 'Asia Region')
        lgech_count = len(df[lgech_condition])
        df.loc[lgech_condition, 'Final Region'] = 'China'
        logger.info(f"    → 변경 완료: {lgech_count}건")

        # 최종 Final Region 고유값 확인
        final_regions = df['Final Region'].unique()
        logger.info(f"  - 최종 Final Region 고유값 ({len(final_regions)}개): {list(final_regions)}")

        # 각 Region별 건수 출력
        logger.info("  - Region별 데이터 건수:")
        # NaN 값 제외하고 정렬
        valid_regions = [r for r in final_regions if pd.notna(r)]

        for region in sorted(valid_regions, key=str):
            count = len(df[df['Final Region'] == region])
            logger.info(f"    {region}: {count}건")

        # NaN 값이 있는 경우 한 번만 출력
        null_count = len(df[df['Final Region'].isna()])
        if null_count > 0:
            logger.info(f"    (null): {null_count}건")

        logger.info("✓ Region(MP) 수정 완료")

        return df

    except Exception as e:
        logger.error(f"✗ 오류 발생: {e}")
        return None

def extract_new_hire_complete(df_input, file_path, detail_sheet, analysis_year):
    """
    hire data에서 올해 입사자를 추출하는 완전한 함수 (6단계)

    Args:
        df_input: 입력 데이터프레임
        file_path: Excel 파일 경로
        detail_sheet: 시트명
        analysis_year: 분석 기준 년도
    """
    try:
        # 6단계: New Hire 추출
        logger.info("6단계: New Hire을 추출합니다...")

        # 데이터 읽기 (df_input을 사용하도록 변경)
        df = df_input.copy()  # 원본 데이터프레임 변경 방지

        # hire date 컬럼 찾기
        hire_date_col = None
        for col in df.columns:
            if 'hire' in col.lower() and 'date' in col.lower():
                hire_date_col = col
                break

        if hire_date_col is None:
            logger.error("✗ hire date 컬럼을 찾을 수 없습니다.")
            return None

        logger.info(f"✓ hire date 컬럼 발견: '{hire_date_col}'")

        # 분석 기준 년도 사용
        current_year = analysis_year
        logger.info(f"  - 분석 기준 년도: {current_year}")

        # New Hire 컬럼 생성
        logger.info("  - 'New Hire' 컬럼 생성 중...")
        logger.info("    New Hire 판단 로직:")
        logger.info(f"      1) hire date가 null이면 → 'N'")
        logger.info(f"      2) hire date의 년도가 {current_year}년이면 → 'Y'")
        logger.info(f"      3) hire date의 년도가 {current_year}년이 아니면 → 'N'")
        logger.info(f"      4) 날짜 변환 실패 시 → 'N'")

        def get_new_hire_status(hire_date):
            if pd.isna(hire_date):
                return 'N'
            try:
                # 날짜를 datetime으로 변환
                if isinstance(hire_date, str):
                    hire_date = pd.to_datetime(hire_date)
                elif hasattr(hire_date, 'year'):
                    # 이미 datetime 객체인 경우
                    pass
                else:
                    return 'N'

                # 년도 비교
                if hire_date.year == current_year:
                    return 'Y'
                else:
                    return 'N'
            except:
                return 'N'

        logger.info("  - New Hire 로직 적용 중...")
        df['New Hire'] = df[hire_date_col].apply(get_new_hire_status)

        # 결과 요약
        new_hire_counts = df['New Hire'].value_counts()
        logger.info("  - New Hire 생성 결과:")
        total_count = len(df)
        for status, count in new_hire_counts.items():
            percentage = (count / total_count * 100) if total_count > 0 else 0
            logger.info(f"    {status}: {count}건 ({percentage:.1f}%)")
        logger.info(f"    전체: {total_count}건")

        logger.info("✓ New Hire 추출 완료")

        return df

    except Exception as e:
        logger.error(f"✗ 오류 발생: {e}")
        return None

def update_branch_mapping(df_input, file_path, detail_sheet):
    """
    Branch 매핑을 처리하는 함수 (7단계)
    """
    try:
        # 7단계: Branch 매핑 처리

        # 데이터 읽기 (df_input을 사용하도록 변경)
        df = df_input.copy()  # 원본 데이터프레임 변경 방지

        # Sub. Name(MP) 컬럼 찾기
        sub_name_col = None
        for col in df.columns:
            if 'sub' in col.lower() and 'name' in col.lower() and 'mp' in col.lower() and 'integrated' not in col.lower():
                sub_name_col = col
                break

        if sub_name_col is None:
            logger.error("✗ Sub. Name(MP) 컬럼을 찾을 수 없습니다.")
            return None

        # Final Sub. 컬럼을 직접 확인하여 매핑
        logger.info("7단계: Final Sub. 직접 수정")
        logger.info("  - 요구사항 반영:")
        logger.info("    1) LGECA → LGECL 변경")
        logger.info("    2) LGEIC → LGERC 변경")
        logger.info("    3) LGECE (Sub. Name(MP)=Czech) → LGECZ 변경")

        # 매핑 전 상태 확인
        lgeca_count = len(df[df['Final Sub.'] == 'LGECA'])
        lgeic_count = len(df[df['Final Sub.'] == 'LGEIC'])
        lgece_total_count = len(df[df['Final Sub.'] == 'LGECE'])
        lgece_czech_count = len(df[(df['Final Sub.'] == 'LGECE') & (df[sub_name_col] == 'Czech')])

        logger.info(f"  - 매핑 전 상태:")
        logger.info(f"    LGECA: {lgeca_count}건")
        logger.info(f"    LGEIC: {lgeic_count}건")
        logger.info(f"    LGECE (전체): {lgece_total_count}건")
        logger.info(f"    LGECE (Czech): {lgece_czech_count}건")

        # Final Sub.에서 직접 매핑
        logger.info("  - 매핑 작업 수행 중...")
        if lgeca_count > 0:
            df.loc[df['Final Sub.'] == 'LGECA', 'Final Sub.'] = 'LGECL'
            logger.info(f"    ✓ LGECA → LGECL 변경 완료: {lgeca_count}건")
        else:
            logger.info(f"    - LGECA 데이터 없음 (변경 불필요)")

        if lgeic_count > 0:
            df.loc[df['Final Sub.'] == 'LGEIC', 'Final Sub.'] = 'LGERC'
            logger.info(f"    ✓ LGEIC → LGERC 변경 완료: {lgeic_count}건")
        else:
            logger.info(f"    - LGEIC 데이터 없음 (변경 불필요)")

        if lgece_czech_count > 0:
            df.loc[(df['Final Sub.'] == 'LGECE') & (df[sub_name_col] == 'Czech'), 'Final Sub.'] = 'LGECZ'
            logger.info(f"    ✓ LGECE (Czech) → LGECZ 변경 완료: {lgece_czech_count}건")
        else:
            logger.info(f"    - LGECE (Czech) 데이터 없음 (변경 불필요)")

        # 매핑 후 최종 상태 확인
        lgectl_count = len(df[df['Final Sub.'] == 'LGECL'])
        lgerc_count = len(df[df['Final Sub.'] == 'LGERC'])
        lgece_remaining_count = len(df[df['Final Sub.'] == 'LGECE'])
        lgecz_count = len(df[df['Final Sub.'] == 'LGECZ'])

        logger.info(f"  - 매핑 후 상태:")
        logger.info(f"    LGECL: {lgectl_count}건 (변경 전 LGECA: {lgeca_count}건)")
        logger.info(f"    LGERC: {lgerc_count}건 (변경 전 LGEIC: {lgeic_count}건)")
        logger.info(f"    LGECE (전체): {lgece_total_count}건 → LGECE (남은): {lgece_remaining_count}건, LGECZ (신규): {lgecz_count}건")
        logger.info(f"      (Czech만 LGECZ로 변경: {lgece_czech_count}건)")

        logger.info("✓ 7단계 완료: Final Sub. 직접 수정 완료")

        return df

    except Exception as e:
        logger.error(f"✗ 오류 발생: {e}")
        return None

def create_new_leader_column(df_input, file_directory, analysis_year):
    """
    New Leader 컬럼을 생성하는 함수 (7단계)
    prev_hr_index_final.csv와 조인하여 신임 팀장을 식별

    Args:
        df_input: 입력 데이터프레임 (hr_index_final)
        file_directory: 파일 디렉토리 경로
        analysis_year: 분석 기준 년도

    Returns:
        DataFrame: New Leader 컬럼이 추가된 데이터프레임
    """
    logger.info("7단계: New Leader 컬럼 생성")
    logger.info("=" * 80)

    try:
        df = df_input.copy()

        # prev_hr_index_final.csv 파일 로드
        prev_hr_path = os.path.join(file_directory, 'prev_hr_index_final.csv')

        if not os.path.exists(prev_hr_path):
            logger.warning(f"✗ prev_hr_index_final.csv 파일을 찾을 수 없습니다: {prev_hr_path}")
            logger.warning("  - New Leader 컬럼을 모두 'N'으로 설정합니다.")
            df['New Leader'] = 'N'
            df['Position_prev'] = None
            return df

        logger.info(f"  - 사용 파일: {prev_hr_path}")
        prev_df = pd.read_csv(prev_hr_path)
        logger.info(f"  - prev_hr_index_final.csv 로드 완료: {len(prev_df)}행")

        # 조인 전 상태
        before_count = len(df)
        logger.info(f"  - 조인 전 hr_index_final: {before_count}행")

        # Emp. No. 컬럼 확인
        emp_col_current = None
        for col in df.columns:
            if col.lower() in ['emp. no.', 'emp no', 'emp.no.', 'employee number', 'empno']:
                emp_col_current = col
                break

        emp_col_prev = None
        for col in prev_df.columns:
            if col.lower() in ['emp. no.', 'emp no', 'emp.no.', 'employee number', 'empno']:
                emp_col_prev = col
                break

        if emp_col_current is None or emp_col_prev is None:
            logger.error(f"✗ Emp. No. 컬럼을 찾을 수 없습니다.")
            logger.error(f"  - 현재 HR: {emp_col_current}")
            logger.error(f"  - 이전 HR: {emp_col_prev}")
            df['New Leader'] = 'N'
            df['Position_prev'] = None
            return df

        logger.info(f"  - 조인 키 컬럼: '{emp_col_current}' (현재) ← '{emp_col_prev}' (이전)")

        # Position 컬럼 확인
        position_col_prev = None
        for col in prev_df.columns:
            if col.lower() == 'position':
                position_col_prev = col
                break

        if position_col_prev is None:
            logger.error(f"✗ prev_hr_index_final에서 Position 컬럼을 찾을 수 없습니다.")
            df['New Leader'] = 'N'
            df['Position_prev'] = None
            return df

        # prev_df에서 Emp. No.와 Position만 선택하고 rename
        prev_df_subset = prev_df[[emp_col_prev, position_col_prev]].copy()
        prev_df_subset = prev_df_subset.rename(columns={position_col_prev: 'Position_prev'})

        logger.info(f"  - prev_hr에서 가져올 컬럼: Position → Position_prev")

        # LEFT JOIN 수행
        logger.info("  - LEFT JOIN 수행 중...")
        df_merged = df.merge(
            prev_df_subset,
            left_on=emp_col_current,
            right_on=emp_col_prev,
            how='left',
            suffixes=('', '_drop')
        )

        # 중복 컬럼 제거 (조인 키)
        if emp_col_prev in df_merged.columns and emp_col_prev != emp_col_current:
            df_merged = df_merged.drop(columns=[emp_col_prev])

        logger.info(f"  ✓ LEFT JOIN 완료")
        logger.info(f"    조인 후: {len(df_merged)}행")

        # 매칭 통계
        matched_count = df_merged['Position_prev'].notna().sum()
        unmatched_count = df_merged['Position_prev'].isna().sum()
        logger.info(f"    매칭 성공: {matched_count}건")
        logger.info(f"    매칭 실패 (신규 직원 등): {unmatched_count}건")
        logger.info("")

        # New Leader 조건 확인
        logger.info("7.1: New Leader 판단 조건")
        logger.info(f"  - 조건 1: FSE_ISE = 'ISE' AND Position in ['Team Leader', 'Leader_팀장'] AND Position_prev NOT IN ['Team Leader', 'Leader_팀장']")
        logger.info(f"  - 조건 2: FSE_ISE = 'ISE' AND Position in ['Team Leader', 'Leader_팀장'] AND Hire Date가 {analysis_year}년")
        logger.info("  - 위 조건 중 하나라도 만족하면 'Y', 아니면 'N'")
        logger.info("")

        # 필요한 컬럼 찾기
        # FSE_ISE 컬럼
        fse_ise_col = None
        for col in df_merged.columns:
            if col.lower() in ['fse_ise', 'fse/ise']:
                fse_ise_col = col
                break

        # Position 컬럼
        position_col = None
        for col in df_merged.columns:
            if col.lower() == 'position':
                position_col = col
                break

        # Hire Date 컬럼
        hire_date_col = None
        for col in df_merged.columns:
            if 'hire' in col.lower() and 'date' in col.lower():
                hire_date_col = col
                break

        if fse_ise_col is None or position_col is None or hire_date_col is None:
            logger.error(f"✗ 필수 컬럼을 찾을 수 없습니다:")
            logger.error(f"    FSE_ISE: {fse_ise_col}")
            logger.error(f"    Position: {position_col}")
            logger.error(f"    Hire Date: {hire_date_col}")
            df_merged['New Leader'] = 'N'
            return df_merged

        logger.info(f"  - 사용 컬럼:")
        logger.info(f"    FSE_ISE: '{fse_ise_col}'")
        logger.info(f"    Position: '{position_col}'")
        logger.info(f"    Position_prev: 'Position_prev'")
        logger.info(f"    Hire Date: '{hire_date_col}'")
        logger.info("")

        # New Leader 계산 함수
        def get_new_leader_status(row):
            """각 행의 New Leader 상태를 판단"""
            # FSE_ISE 체크
            fse_ise = row.get(fse_ise_col)
            if pd.isna(fse_ise) or str(fse_ise).strip().upper() != 'ISE':
                return 'N'

            # Position 체크
            position = row.get(position_col)
            if pd.isna(position):
                return 'N'

            position_str = str(position).strip()
            is_leader = position_str in ['Team Leader', 'Leader_팀장']

            if not is_leader:
                return 'N'

            # 여기까지 왔다면: FSE_ISE = 'ISE' AND Position in ['Team Leader', 'Leader_팀장']

            # 조건 1: Position_prev가 Team Leader/Leader_팀장이 아닌 경우
            position_prev = row.get('Position_prev')
            if pd.notna(position_prev):
                position_prev_str = str(position_prev).strip()
                if position_prev_str not in ['Team Leader', 'Leader_팀장']:
                    return 'Y'  # 승진한 경우
            else:
                # Position_prev가 null이면 이전 데이터 없음 → Team Leader가 아닌 것으로 간주
                return 'Y'

            # 조건 2: Hire Date가 분석 년도인 경우
            hire_date = row.get(hire_date_col)
            if pd.notna(hire_date):
                try:
                    if isinstance(hire_date, str):
                        hire_dt = pd.to_datetime(hire_date)
                    elif hasattr(hire_date, 'year'):
                        hire_dt = hire_date
                    else:
                        return 'N'

                    if hire_dt.year == analysis_year:
                        return 'Y'  # 올해 고용된 팀장
                except:
                    pass

            return 'N'

        logger.info("7.2: New Leader 계산 중...")
        df_merged['New Leader'] = df_merged.apply(get_new_leader_status, axis=1)

        # 통계 출력
        new_leader_y_count = (df_merged['New Leader'] == 'Y').sum()
        new_leader_n_count = (df_merged['New Leader'] == 'N').sum()

        logger.info(f"  ✓ New Leader 계산 완료")
        logger.info(f"    New Leader = 'Y': {new_leader_y_count}명 ({new_leader_y_count/len(df_merged)*100:.2f}%)")
        logger.info(f"    New Leader = 'N': {new_leader_n_count}명 ({new_leader_n_count/len(df_merged)*100:.2f}%)")
        logger.info("")

        # 조건별 상세 분석
        logger.info("7.3: New Leader 상세 분석")

        # FSE_ISE = 'ISE'인 사람 수
        ise_count = (df_merged[fse_ise_col].astype(str).str.strip().str.upper() == 'ISE').sum()
        logger.info(f"  - FSE_ISE = 'ISE': {ise_count}명")

        # Position이 Team Leader 또는 Leader_팀장인 사람 수
        is_leader_mask = df_merged[position_col].astype(str).str.strip().isin(['Team Leader', 'Leader_팀장'])
        leader_count = is_leader_mask.sum()
        logger.info(f"  - Position in ['Team Leader', 'Leader_팀장']: {leader_count}명")

        # FSE_ISE = 'ISE' AND Position in ['Team Leader', 'Leader_팀장']
        ise_leader_mask = (df_merged[fse_ise_col].astype(str).str.strip().str.upper() == 'ISE') & is_leader_mask
        ise_leader_count = ise_leader_mask.sum()
        logger.info(f"  - FSE_ISE = 'ISE' AND Position in ['Team Leader', 'Leader_팀장']: {ise_leader_count}명")
        logger.info("")

        # 조건 1: 승진한 경우
        condition1_mask = ise_leader_mask & (
            (df_merged['Position_prev'].isna()) |
            (~df_merged['Position_prev'].astype(str).str.strip().isin(['Team Leader', 'Leader_팀장']))
        )
        condition1_count = condition1_mask.sum()
        logger.info(f"  - 조건 1 (승진): {condition1_count}명")
        logger.info(f"    FSE_ISE='ISE' AND 현재 Team Leader AND 이전 NOT Team Leader")

        if condition1_count > 0 and condition1_count <= 10:
            condition1_subs = df_merged[condition1_mask][[emp_col_current, position_col, 'Position_prev']].head(10)
            logger.info(f"    샘플 (최대 10명):")
            for idx, row in condition1_subs.iterrows():
                prev_pos = row['Position_prev'] if pd.notna(row['Position_prev']) else 'null'
                logger.info(f"      Emp: {row[emp_col_current]}, Position: {row[position_col]}, Position_prev: {prev_pos}")

        # 조건 2: 올해 고용된 팀장
        def is_hired_this_year(hire_date):
            if pd.isna(hire_date):
                return False
            try:
                if isinstance(hire_date, str):
                    hire_dt = pd.to_datetime(hire_date)
                elif hasattr(hire_date, 'year'):
                    hire_dt = hire_date
                else:
                    return False
                return hire_dt.year == analysis_year
            except:
                return False

        condition2_mask = ise_leader_mask & df_merged[hire_date_col].apply(is_hired_this_year)
        condition2_count = condition2_mask.sum()
        logger.info("")
        logger.info(f"  - 조건 2 (올해 고용): {condition2_count}명")
        logger.info(f"    FSE_ISE='ISE' AND 현재 Team Leader AND Hire Date = {analysis_year}년")

        if condition2_count > 0 and condition2_count <= 10:
            condition2_subs = df_merged[condition2_mask][[emp_col_current, hire_date_col, position_col]].head(10)
            logger.info(f"    샘플 (최대 10명):")
            for idx, row in condition2_subs.iterrows():
                logger.info(f"      Emp: {row[emp_col_current]}, Hire Date: {row[hire_date_col]}, Position: {row[position_col]}")

        # 중복 체크 (조건 1과 조건 2를 동시에 만족하는 경우)
        both_conditions = condition1_mask & condition2_mask
        both_count = both_conditions.sum()
        logger.info("")
        logger.info(f"  - 조건 1 AND 조건 2 동시 만족: {both_count}명")
        logger.info("")

        logger.info(f"✓ 7단계 완료: New Leader 컬럼 생성 완료")
        logger.info(f"  - New Leader = 'Y': {new_leader_y_count}명")
        logger.info(f"  - New Leader = 'N': {new_leader_n_count}명")
        logger.info(f"  - Position_prev 컬럼 추가 완료 (매칭: {matched_count}명, 미매칭: {unmatched_count}명)")
        logger.info("=" * 80)
        logger.info("")

        return df_merged

    except Exception as e:
        logger.error(f"✗ New Leader 컬럼 생성 중 오류 발생: {e}")
        import traceback
        logger.error(f"상세 오류:\n{traceback.format_exc()}")
        df['New Leader'] = 'N'
        df['Position_prev'] = None
        return df

def filter_manage_area(df_input, file_directory, index_management_file):
    """
    Manage Area 필터링 함수 (8단계)
    index_management_final.csv와 조인하여 Manage Area = 'N'인 레코드 제거

    Args:
        df_input (pandas.DataFrame): 입력 데이터프레임
        file_directory (str): 파일 디렉토리 경로
        index_management_file (str): index_management_final.csv 파일명

    Returns:
        pandas.DataFrame: 필터링된 데이터프레임
    """
    try:
        logger.info("8단계: Manage Area 필터링 (관리하지 않는 법인 제거)")
        logger.info("  - 작업 내용:")
        logger.info("    1) index_management_final.csv 파일 읽기")
        logger.info("    2) Final Sub. 기준으로 LEFT JOIN")
        logger.info("    3) Manage Area = 'N'인 레코드 삭제")
        logger.info("  - 목적: 관리 대상이 아닌 법인 데이터 제거")

        df = df_input.copy()

        # index_management_final.csv 파일 읽기
        index_mgmt_path = os.path.join(file_directory, index_management_file)

        if not os.path.exists(index_mgmt_path):
            logger.warning(f"✗ index_management_final.csv 파일을 찾을 수 없습니다: {index_mgmt_path}")
            logger.info("    필터링 없이 진행합니다.")
            return df

        logger.info(f"  - Index Management 파일 로드: {index_mgmt_path}")
        index_mgmt_df = pd.read_csv(index_mgmt_path, encoding='utf-8-sig')
        logger.info(f"    ✓ 로드 완료: {len(index_mgmt_df)}행, {len(index_mgmt_df.columns)}열")

        # 필요한 컬럼 확인
        if 'Final Sub.' not in index_mgmt_df.columns or 'Manage Area' not in index_mgmt_df.columns:
            logger.warning("✗ index_management_final.csv에서 필요한 컬럼을 찾을 수 없습니다.")
            logger.info(f"    사용 가능한 컬럼: {list(index_mgmt_df.columns)}")
            logger.info("    필터링 없이 진행합니다.")
            return df

        logger.info(f"    ✓ 필요 컬럼 확인: Final Sub., Manage Area")

        # Manage Area 분포 확인
        manage_area_dist = index_mgmt_df['Manage Area'].value_counts()
        logger.info(f"  - Manage Area 분포 (index_management_final.csv):")
        for value, count in manage_area_dist.items():
            logger.info(f"    {value}: {count}개 법인")

        # 필터링 전 상태
        before_count = len(df)
        before_subsidiaries = df['Final Sub.'].nunique()

        logger.info(f"  - 필터링 전: {before_count}행, {before_subsidiaries}개 법인")

        # Final Sub. 기준으로 LEFT JOIN
        logger.info("  - LEFT JOIN 수행 중...")
        logger.info("    조인 키: Final Sub. (대소문자 구분 없음)")

        df['Final_Sub_upper'] = df['Final Sub.'].str.strip().str.upper()
        index_mgmt_df['Final_Sub_upper'] = index_mgmt_df['Final Sub.'].str.strip().str.upper()

        df_merged = df.merge(
            index_mgmt_df[['Final_Sub_upper', 'Manage Area']],
            on='Final_Sub_upper',
            how='left'
        )

        # 매칭 결과 확인
        matched_count = df_merged['Manage Area'].notna().sum()
        unmatched_count = df_merged['Manage Area'].isna().sum()
        logger.info(f"  - 조인 결과:")
        logger.info(f"    매칭 성공: {matched_count}건")
        logger.info(f"    매칭 실패: {unmatched_count}건 (index_management에 없는 법인)")

        if unmatched_count > 0:
            unmatched_subs = df_merged[df_merged['Manage Area'].isna()]['Final Sub.'].unique()
            logger.warning(f"    매칭 실패 법인 목록 ({len(unmatched_subs)}개): {', '.join(unmatched_subs)}")

        # Manage Area = 'N'인 레코드 확인
        manage_n_count = len(df_merged[df_merged['Manage Area'] == 'N'])
        manage_y_count = len(df_merged[df_merged['Manage Area'] == 'Y'])

        logger.info(f"  - Manage Area 분류:")
        logger.info(f"    Y (관리 대상): {manage_y_count}건")
        logger.info(f"    N (관리 제외): {manage_n_count}건")
        logger.info(f"    null (매칭 안됨): {unmatched_count}건 → 유지")

        # Manage Area = 'N'인 법인 목록
        if manage_n_count > 0:
            n_subsidiaries = df_merged[df_merged['Manage Area'] == 'N']['Final Sub.'].unique()
            logger.info(f"  - 삭제 대상 법인 ({len(n_subsidiaries)}개):")
            logger.info(f"    {', '.join(sorted(n_subsidiaries))}")

        # Manage Area != 'N'인 레코드만 유지 (Y 또는 null)
        logger.info("  - 필터링 수행: Manage Area = 'N' 제외")
        df_filtered = df_merged[df_merged['Manage Area'] != 'N'].copy()

        # 임시 컬럼 제거
        df_filtered = df_filtered.drop(['Final_Sub_upper', 'Manage Area'], axis=1, errors='ignore')

        # 필터링 후 상태
        after_count = len(df_filtered)
        after_subsidiaries = df_filtered['Final Sub.'].nunique()
        deleted_count = before_count - after_count
        deleted_subs = before_subsidiaries - after_subsidiaries

        logger.info(f"  - 필터링 후: {after_count}행, {after_subsidiaries}개 법인")
        logger.info(f"  - 삭제된 레코드: {deleted_count}건 ({deleted_count/before_count*100:.1f}%)")
        logger.info(f"  - 삭제된 법인: {deleted_subs}개")
        logger.info(f"✓ 8단계 완료: Manage Area 필터링 완료")

        return df_filtered

    except Exception as e:
        logger.error(f"✗ 오류 발생: {e}")
        return None

def update_region_mp(file_path, detail_sheet):
    """
    Region(MP) 컬럼을 수정하는 함수

    Args:
        file_path (str): Excel 파일 경로
        detail_sheet (str): DETAIL 시트명

    Returns:
        pandas.DataFrame: final_region 컬럼이 추가된 데이터프레임
    """
    try:
        # 5단계: Region(MP) 수정
        logger.info("5단계: Region(MP)을 수정합니다...")

        # 데이터 읽기
        df = pd.read_excel(file_path, sheet_name=detail_sheet)

        # Region(MP) 컬럼 찾기
        region_col = None
        for col in df.columns:
            if 'region' in col.lower() and 'mp' in col.lower():
                region_col = col
                break

        if region_col is None:
            logger.error("✗ Region(MP) 컬럼을 찾을 수 없습니다.")
            return None

        logger.info(f"✓ Region(MP) 컬럼: {region_col}")

        # Integrated Sub. Name(MP) 컬럼 찾기
        integrated_col = None
        for col in df.columns:
            if 'integrated' in col.lower() and 'sub' in col.lower() and 'name' in col.lower() and 'mp' in col.lower():
                integrated_col = col
                break

        if integrated_col is None:
            logger.error("✗ Integrated Sub. Name(MP) 컬럼을 찾을 수 없습니다.")
            return None

        # Final Region 컬럼 생성 (기존 Region(MP) 복사)
        logger.info("  - 'Final Region' 컬럼 생성 중...")
        logger.info(f"    기본값: 'Region(MP)' 컬럼 ({region_col}) 값을 복사")
        df['Final Region'] = df[region_col].copy()

        # 생성된 Final Region의 고유값 확인
        initial_regions = df['Final Region'].unique()
        logger.info(f"    생성된 Final Region 고유값 ({len(initial_regions)}개): {list(initial_regions)}")

        # LGECH + Asia Region → China 변경
        logger.info("  - 특수 케이스 처리: LGECH + Asia Region → China 변경")
        logger.info("    조건: Integrated Sub. Name(MP) = 'LGECH' AND Region(MP) = 'Asia Region'")
        lgech_condition = (df[integrated_col] == 'LGECH') & (df[region_col] == 'Asia Region')
        lgech_count = len(df[lgech_condition])
        df.loc[lgech_condition, 'Final Region'] = 'China'
        logger.info(f"    → 변경 완료: {lgech_count}건")

        # 최종 Final Region 고유값 확인
        final_regions = df['Final Region'].unique()
        logger.info(f"  - 최종 Final Region 고유값 ({len(final_regions)}개): {list(final_regions)}")

        # 각 Region별 건수 출력
        logger.info("  - Region별 데이터 건수:")
        # NaN 값 제외하고 정렬
        valid_regions = [r for r in final_regions if pd.notna(r)]

        for region in sorted(valid_regions, key=str):
            count = len(df[df['Final Region'] == region])
            logger.info(f"    {region}: {count}건")

        # NaN 값이 있는 경우 한 번만 출력
        null_count = len(df[df['Final Region'].isna()])
        if null_count > 0:
            logger.info(f"    (null): {null_count}건")

        logger.info("✓ Region(MP) 수정 완료")

        return df

    except Exception as e:
        logger.error(f"✗ 오류 발생: {e}")
        return None

def extract_new_hire(file_path, detail_sheet):
    """
    hire data에서 올해 입사자를 추출하는 함수

    Args:
        file_path (str): Excel 파일 경로
        detail_sheet (str): DETAIL 시트명

    Returns:
        pandas.DataFrame: new_hire 컬럼이 추가된 데이터프레임
    """
    try:
        # 6단계: New Hire 추출
        logger.info("6단계: New Hire을 추출합니다...")

        # 데이터 읽기
        df = pd.read_excel(file_path, sheet_name=detail_sheet)

        # hire date 컬럼 찾기
        hire_date_col = None
        for col in df.columns:
            if 'hire' in col.lower() and 'date' in col.lower():
                hire_date_col = col
                break

        if hire_date_col is None:
            logger.error("✗ hire date 컬럼을 찾을 수 없습니다.")
            return None

        logger.info(f"✓ hire date 컬럼: {hire_date_col}")

        # 현재 년도 가져오기
        current_year = pd.Timestamp.now().year
        logger.info(f"✓ 현재 년도: {current_year}")

        # new_hire 컬럼 생성
        def get_new_hire_status(hire_date):
            if pd.isna(hire_date):
                return 'N'
            try:
                # 날짜를 datetime으로 변환
                if isinstance(hire_date, str):
                    hire_date = pd.to_datetime(hire_date)
                elif hasattr(hire_date, 'year'):
                    # 이미 datetime 객체인 경우
                    pass
                else:
                    return 'N'

                # 년도 비교
                if hire_date.year == current_year:
                    return 'Y'
                else:
                    return 'N'
            except:
                return 'N'

        df['New Hire'] = df[hire_date_col].apply(get_new_hire_status)

        # 결과 요약
        new_hire_counts = df['New Hire'].value_counts()
        logger.info("New Hire 분포:")
        for status, count in new_hire_counts.items():
            logger.info(f"  {status}: {count}개")

        logger.info("✓ New Hire 추출 완료")

        return df

    except Exception as e:
        logger.error(f"✗ 오류 발생: {e}")
        return None
