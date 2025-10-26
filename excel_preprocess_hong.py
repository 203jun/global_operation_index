#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HONG Data Excel 파일 전처리 스크립트
HONG Data Excel 파일에서 법인담당자 시트의 컬럼 리스트를 추출하는 기능을 제공합니다.
"""

import pandas as pd
import os
import sys
from logger_config import get_default_logger

# 로거 설정
logger = get_default_logger(__name__)

def load_hong_excel_file(file_path):
    """
    HONG Data Excel 파일을 불러오는 함수
    """
    logger.info("1단계: hong_data.xlsx 파일을 불러옵니다...")
    xl_file = pd.ExcelFile(file_path)
    logger.info("✓ Excel 파일 불러오기 완료")
    return xl_file

def find_legal_person_sheet(xl_file):
    """
    법인담당자 시트를 찾는 함수
    """
    logger.info("2단계: 법인담당자 시트를 찾습니다...")
    legal_person_sheet = None
    for sheet_name in xl_file.sheet_names:
        if '법인담당자' in sheet_name:
            legal_person_sheet = sheet_name
            break

    if legal_person_sheet is None:
        logger.error("✗ 법인담당자 시트를 찾을 수 없습니다.")
        return None

    logger.info(f"✓ 법인담당자 시트 발견: {legal_person_sheet}")
    return legal_person_sheet

def find_annual_plan_sheet(xl_file):
    """
    연간교육계획 시트를 찾는 함수
    """
    logger.info("2단계: 연간교육계획 시트를 찾습니다...")
    annual_plan_sheet = None
    for sheet_name in xl_file.sheet_names:
        if '연간교육계획' in sheet_name:
            annual_plan_sheet = sheet_name
            break

    if annual_plan_sheet is None:
        logger.error("✗ 연간교육계획 시트를 찾을 수 없습니다.")
        return None

    logger.info(f"✓ 연간교육계획 시트 발견: {annual_plan_sheet}")
    return annual_plan_sheet

def get_legal_person_columns(file_path, legal_person_sheet):
    """
    Excel 파일에서 법인담당자 시트의 모든 컬럼명을 반환하는 함수
    """
    try:
        # 3단계: 컬럼명 확인
        logger.info("3단계: 법인담당자 시트의 컬럼명을 확인합니다...")
        df = pd.read_excel(file_path, sheet_name=legal_person_sheet, nrows=0)
        columns = list(df.columns)

        logger.info(f"✓ 컬럼명 확인 완료: 총 {len(columns)}개 컬럼")
        logger.info(f"컬럼 리스트: {columns}")

        return columns

    except Exception as e:
        logger.error(f"✗ 오류 발생: {e}")
        return None

def get_annual_plan_columns(file_path, annual_plan_sheet):
    """
    Excel 파일에서 연간교육계획 시트의 모든 컬럼명을 반환하는 함수
    """
    try:
        # 3단계: 컬럼명 확인
        logger.info("3단계: 연간교육계획 시트의 컬럼명을 확인합니다...")
        df = pd.read_excel(file_path, sheet_name=annual_plan_sheet, nrows=0)
        columns = list(df.columns)

        logger.info(f"✓ 컬럼명 확인 완료: 총 {len(columns)}개 컬럼")
        logger.info(f"컬럼 리스트: {columns}")

        return columns

    except Exception as e:
        logger.error(f"✗ 오류 발생: {e}")
        return None

def update_subsidiary_mapping(df):
    """
    Subsidiary 값을 수정하는 함수
    """
    logger.info("5단계: Subsidiary 값을 수정합니다...")
    logger.info("  목적: hr_index_final.csv와 조인하기 위해 Subsidiary 값을 통일")
    logger.info("  이유: 두 파일의 Subsidiary 값이 다르면 조인 시 매칭되지 않음")
    logger.info("  방법: hong_data.xlsx의 Subsidiary 값을 hr_index_final.csv의 'Final Sub.' 값에 맞춤")

    # Subsidiary 컬럼 찾기
    subsidiary_col = None
    for col in df.columns:
        if 'subsidiary' in col.lower():
            subsidiary_col = col
            break

    if subsidiary_col is None:
        logger.error("✗ Subsidiary 컬럼을 찾을 수 없습니다.")
        return None

    logger.info(f"✓ Subsidiary 컬럼 발견: '{subsidiary_col}'")

    # 수정 전 고유값 확인
    original_values = df[subsidiary_col].unique()
    logger.info(f"  - 수정 전 Subsidiary 고유값 ({len(original_values)}개): {list(original_values)}")

    # 매핑 정의
    subsidiary_mappings = {
        'LGEIH(Greece Branch)': 'Greece Branch',
        'LGESI': 'LGSI',
        'LGESP-MAO': 'LGESP Manaus Factory',
        'LGEIS': 'LGEIH',
        'LGEIL Noida': 'LGEIL Noida Factory',
        'LGETU': 'Tunisia Branch',
        'LGEYK': 'Israel Branch',
        'LGEHQ': 'Europe Region',
        'Europe RHQ': 'Europe Region',
        'LGEAG': 'Austria Branch',
        'LGEHS': 'Greece Branch',
        'LGELA': 'Latvia Branch',
        'LGEPT': 'Portugal Branch',
        'LGERO': 'Romania Branch',
        'LGECS': 'Central South EU Branch'
    }

    logger.info("  - Subsidiary 매핑 규칙 적용 중...")
    total_changed = 0

    # 각 매핑에 대해 처리
    for old_name, new_name in subsidiary_mappings.items():
        condition = df[subsidiary_col] == old_name
        count = len(df[condition])
        if count > 0:
            df.loc[condition, subsidiary_col] = new_name
            logger.info(f"    '{old_name}' → '{new_name}': {count}건")
            total_changed += count

    # 수정 후 고유값 확인
    final_values = df[subsidiary_col].unique()
    logger.info(f"  - 수정 후 Subsidiary 고유값 ({len(final_values)}개): {list(final_values)}")
    logger.info(f"  - 총 변경된 건수: {total_changed}건 / 전체 {len(df)}건")

    # Subsidiary 대문자 변환
    logger.info("  - Subsidiary를 대문자로 변환합니다...")
    logger.info("    목적: hr_index의 Final Sub.와 조인 시 키 매칭을 정확하게 하기 위함")
    logger.info("    이유: 조인 규칙 - 모든 법인명은 대문자로 통일")
    logger.info("    예시: 'lgejp' → 'LGEJP', 'Europe Region' → 'EUROPE REGION'")

    # 변환 전 샘플
    before_sample = df[subsidiary_col].head(5).tolist()
    logger.info(f"    변환 전 샘플: {before_sample}")

    # 대문자 변환
    df[subsidiary_col] = df[subsidiary_col].str.upper()

    # 변환 후 샘플
    after_sample = df[subsidiary_col].head(5).tolist()
    logger.info(f"    변환 후 샘플: {after_sample}")

    # 변환 후 고유값 확인
    upper_values = df[subsidiary_col].unique()
    logger.info(f"    변환 후 Subsidiary 고유값 ({len(upper_values)}개): {list(upper_values)}")
    logger.info("    ✓ Subsidiary 대문자 변환 완료")

    logger.info("✓ 5단계 완료: Subsidiary 매핑 처리 완료")
    return df

def filter_hong_manage_area(df_input, file_directory, index_management_file):
    """
    Manage Area 필터링 함수 (6단계)
    index_management_final.csv와 조인하여 Manage Area = 'N'인 레코드 제거

    Args:
        df_input (pandas.DataFrame): 입력 데이터프레임
        file_directory (str): 파일 디렉토리 경로
        index_management_file (str): index_management_final.csv 파일명

    Returns:
        pandas.DataFrame: 필터링된 데이터프레임
    """
    try:
        logger.info("6단계: Manage Area 필터링 (관리하지 않는 법인 제거)")
        logger.info("  - 작업 내용:")
        logger.info("    1) index_management_final.csv 파일 읽기")
        logger.info("    2) Subsidiary = Final Sub. 기준으로 LEFT JOIN")
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
        before_subsidiaries = df['Subsidiary'].nunique()

        logger.info(f"  - 필터링 전: {before_count}행, {before_subsidiaries}개 법인")

        # Subsidiary 기준으로 LEFT JOIN
        logger.info("  - LEFT JOIN 수행 중...")
        logger.info("    조인 키: Subsidiary = Final Sub. (대소문자 구분 없음)")

        df['Subsidiary_upper'] = df['Subsidiary'].str.strip().str.upper()
        index_mgmt_df['Final_Sub_upper'] = index_mgmt_df['Final Sub.'].str.strip().str.upper()

        df_merged = df.merge(
            index_mgmt_df[['Final_Sub_upper', 'Manage Area']],
            left_on='Subsidiary_upper',
            right_on='Final_Sub_upper',
            how='left'
        )

        # 매칭 결과 확인
        matched_count = df_merged['Manage Area'].notna().sum()
        unmatched_count = df_merged['Manage Area'].isna().sum()
        logger.info(f"  - 조인 결과:")
        logger.info(f"    매칭 성공: {matched_count}건")
        logger.info(f"    매칭 실패: {unmatched_count}건 (index_management에 없는 법인)")

        if unmatched_count > 0:
            unmatched_subs = df_merged[df_merged['Manage Area'].isna()]['Subsidiary'].unique()
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
            n_subsidiaries = df_merged[df_merged['Manage Area'] == 'N']['Subsidiary'].unique()
            logger.info(f"  - 삭제 대상 법인 ({len(n_subsidiaries)}개):")
            logger.info(f"    {', '.join(sorted(n_subsidiaries))}")

        # Manage Area != 'N'인 레코드만 유지 (Y 또는 null)
        logger.info("  - 필터링 수행: Manage Area = 'N' 제외")
        df_filtered = df_merged[df_merged['Manage Area'] != 'N'].copy()

        # 임시 컬럼 제거
        df_filtered = df_filtered.drop(['Subsidiary_upper', 'Final_Sub_upper', 'Manage Area'], axis=1, errors='ignore')

        # 필터링 후 상태
        after_count = len(df_filtered)
        after_subsidiaries = df_filtered['Subsidiary'].nunique()
        deleted_count = before_count - after_count
        deleted_subs = before_subsidiaries - after_subsidiaries

        logger.info(f"  - 필터링 후: {after_count}행, {after_subsidiaries}개 법인")
        logger.info(f"  - 삭제된 레코드: {deleted_count}건 ({deleted_count/before_count*100:.1f}%)")
        logger.info(f"  - 삭제된 법인: {deleted_subs}개")
        logger.info(f"✓ 6단계 완료: Manage Area 필터링 완료")

        return df_filtered

    except Exception as e:
        logger.error(f"✗ 오류 발생: {e}")
        return None

def add_date_columns(file_path, annual_plan_sheet):
    """
    연간교육계획 시트에 날짜 컬럼을 추가하는 함수
    """
    try:
        # 4단계: 데이터 읽기 및 날짜 컬럼 추가
        logger.info("4단계: 데이터를 읽고 날짜 컬럼을 추가합니다...")
        df = pd.read_excel(file_path, sheet_name=annual_plan_sheet)
        logger.info(f"✓ 데이터 읽기 완료: {df.shape[0]}행, {df.shape[1]}열")

        # 현재 년도 가져오기
        current_year = pd.Timestamp.now().year
        logger.info(f"✓ 현재 년도: {current_year}")

        # Month_start, Date_start 컬럼 찾기
        month_start_col = None
        date_start_col = None
        month_end_col = None
        date_end_col = None

        for col in df.columns:
            if 'month' in col.lower() and 'start' in col.lower():
                month_start_col = col
            elif 'date' in col.lower() and 'start' in col.lower():
                date_start_col = col
            elif 'month' in col.lower() and 'end' in col.lower():
                month_end_col = col
            elif 'date' in col.lower() and 'end' in col.lower():
                date_end_col = col

        if month_start_col is None or date_start_col is None or month_end_col is None or date_end_col is None:
            logger.error("✗ 필요한 날짜 컬럼을 찾을 수 없습니다.")
            logger.error(f"Month_start: {month_start_col}, Date_start: {date_start_col}")
            logger.error(f"Month_end: {month_end_col}, Date_end: {date_end_col}")
            return None

        logger.info(f"✓ 날짜 관련 컬럼 발견:")
        logger.info(f"  - Month_start 컬럼: '{month_start_col}'")
        logger.info(f"  - Date_start 컬럼: '{date_start_col}'")
        logger.info(f"  - Month_end 컬럼: '{month_end_col}'")
        logger.info(f"  - Date_end 컬럼: '{date_end_col}'")

        # Start Date 컬럼 생성 (YYYYMMDD 형태)
        logger.info("  - 'Start Date' 컬럼 생성 중...")
        logger.info("    생성 로직:")
        logger.info(f"      1) Month_start와 Date_start 값을 사용")
        logger.info(f"      2) 형식: {current_year}MMDD (YYYYMMDD)")
        logger.info(f"      3) 값이 null이면 기본값: {current_year}0101")
        logger.info(f"      4) 변환 실패 시 기본값: {current_year}0101")

        def create_start_date(row):
            try:
                month = int(row[month_start_col]) if pd.notna(row[month_start_col]) else 1
                date = int(row[date_start_col]) if pd.notna(row[date_start_col]) else 1
                return f"{current_year}{month:02d}{date:02d}"
            except:
                return f"{current_year}0101"

        # End Date 컬럼 생성 (YYYYMMDD 형태)
        logger.info("  - 'End Date' 컬럼 생성 중...")
        logger.info("    생성 로직:")
        logger.info(f"      1) Month_end와 Date_end 값을 사용")
        logger.info(f"      2) 형식: {current_year}MMDD (YYYYMMDD)")
        logger.info(f"      3) 값이 null이면 기본값: {current_year}1231")
        logger.info(f"      4) 변환 실패 시 기본값: {current_year}1231")

        def create_end_date(row):
            try:
                month = int(row[month_end_col]) if pd.notna(row[month_end_col]) else 12
                date = int(row[date_end_col]) if pd.notna(row[date_end_col]) else 31
                return f"{current_year}{month:02d}{date:02d}"
            except:
                return f"{current_year}1231"

        df['Start Date'] = df.apply(create_start_date, axis=1)
        df['End Date'] = df.apply(create_end_date, axis=1)

        logger.info("✓ Start Date, End Date 컬럼 생성 완료")
        logger.info(f"  - Start Date 샘플 (처음 3개): {df['Start Date'].head(3).tolist()}")
        logger.info(f"  - End Date 샘플 (처음 3개): {df['End Date'].head(3).tolist()}")

        return df

    except Exception as e:
        logger.error(f"✗ 오류 발생: {e}")
        return None

def run_hong_manager_preprocessing(file_directory, hong_file_name, hr_file_name, output_file_name):
    """
    HONG Data Excel 파일 - 법인담당자 시트 전처리 실행

    Args:
        file_directory (str): 파일이 있는 디렉토리
        hong_file_name (str): HONG 파일명
        hr_file_name (str): HR 최종 CSV 파일명
        output_file_name (str): 출력 파일명
    """
    logger.info("=== HONG Data Excel 파일 - 법인담당자 시트 전처리 시작 ===")

    # 파일 경로 설정
    file_path = os.path.join(file_directory, hong_file_name)

    # 파일 존재 확인
    if not os.path.exists(file_path):
        logger.error(f"✗ 파일을 찾을 수 없습니다: {file_path}")
        logger.error("파일 경로를 확인해주세요.")
        return False

    logger.info(f"처리할 파일: {file_path}")

    # 1,2단계: Excel 파일 불러오기 및 법인담당자 시트 찾기
    xl_file = load_hong_excel_file(file_path)
    legal_person_sheet = find_legal_person_sheet(xl_file)

    if legal_person_sheet is None:
        logger.error("✗ 법인담당자 시트를 찾을 수 없습니다.")
        return False

    # 3단계: 법인담당자 시트의 컬럼명 추출
    columns = get_legal_person_columns(file_path, legal_person_sheet)

    if columns is not None:
        # 4단계: 데이터 읽기
        logger.info("4단계: 데이터를 읽습니다...")
        df = pd.read_excel(file_path, sheet_name=legal_person_sheet)
        logger.info(f"✓ 데이터 읽기 완료: {df.shape[0]}행, {df.shape[1]}열")

        # 5단계: hr_index_final.csv와 이메일 매칭하여 Final Sub. 추가
        logger.info("5단계: hr_index_final.csv와 이메일 매칭하여 Final Sub. 추가합니다...")

        try:
            # hr_index_final.csv 로드
            hr_index_path = os.path.join(file_directory, hr_file_name)
            if os.path.exists(hr_index_path):
                hr_df = pd.read_csv(hr_index_path, encoding='utf-8-sig')
                logger.info(f"✓ hr_index_final.csv 로드 완료: {hr_df.shape[0]}행")

                # L&D PIC e-mail 컬럼이 있는지 확인
                if 'L&D PIC e-mail' in df.columns and 'E-Mail Adress' in hr_df.columns:
                    logger.info("✓ 이메일 매칭을 시작합니다...")

                    # 이메일 매칭을 위한 딕셔너리 생성
                    email_to_final_sub = {}
                    for _, row in hr_df.iterrows():
                        email = row['E-Mail Adress']
                        final_sub = row['Final Sub.']
                        if pd.notna(email) and pd.notna(final_sub):
                            email_to_final_sub[email] = final_sub

                    logger.info(f"✓ 이메일 매핑 테이블 생성: {len(email_to_final_sub)}개")

                    # Final Sub. 컬럼 추가 및 매칭
                    df['Final Sub.'] = None
                    matched_count = 0

                    for idx, row in df.iterrows():
                        email = row['L&D PIC e-mail']
                        if pd.notna(email) and email in email_to_final_sub:
                            df.at[idx, 'Final Sub.'] = email_to_final_sub[email]
                            matched_count += 1

                    # 매칭 통계
                    total_count = len(df)
                    not_matched_count = total_count - matched_count
                    matched_percentage = (matched_count / total_count * 100) if total_count > 0 else 0
                    not_matched_percentage = (not_matched_count / total_count * 100) if total_count > 0 else 0

                    logger.info(f"✓ 이메일 매칭 결과:")
                    logger.info(f"  전체: {total_count}건")
                    logger.info(f"  매칭 성공: {matched_count}건 ({matched_percentage:.1f}%)")
                    logger.info(f"  매칭 실패: {not_matched_count}건 ({not_matched_percentage:.1f}%)")

                    # 매칭 결과 샘플 출력
                    matched_data = df[df['Final Sub.'].notna()]
                    if not matched_data.empty:
                        logger.info("  매칭 성공 샘플 (최대 3개):")
                        for idx, row in matched_data.head(3).iterrows():
                            logger.info(f"    이메일: {row['L&D PIC e-mail']} → Final Sub.: {row['Final Sub.']}")

                    # 매칭 실패 샘플 출력
                    not_matched_data = df[df['Final Sub.'].isna()]
                    if not not_matched_data.empty:
                        logger.info("  매칭 실패 샘플 (최대 3개):")
                        for idx, row in not_matched_data.head(3).iterrows():
                            logger.info(f"    이메일: {row['L&D PIC e-mail']} → 매칭 안됨")
                else:
                    logger.warning("✗ 필요한 컬럼이 없습니다. (L&D PIC e-mail 또는 E-Mail Adress)")
            else:
                logger.warning("✗ hr_index_final.csv 파일이 없습니다.")

        except Exception as e:
            logger.error(f"✗ 이메일 매칭 중 오류 발생: {e}")

        # 최종단계: CSV 파일로 저장
        logger.info("최종단계: 최종 결과를 CSV 파일로 저장합니다...")
        output_path = os.path.join(file_directory, output_file_name)
        df.to_csv(output_path, index=False, encoding='utf-8-sig')
        logger.info(f"✓ CSV 파일 저장 완료: {output_path}")
        logger.info(f"✓ 저장된 데이터: {df.shape[0]}행, {df.shape[1]}열")
        logger.info("=== HONG 법인담당자 전처리 완료 ===")
        return True
    else:
        logger.error("✗ 법인담당자 시트 컬럼 추출에 실패했습니다.")
        return False

def run_hong_plan_preprocessing(file_directory, hong_file_name, output_file_name):
    """
    HONG Data Excel 파일 - 연간교육계획 시트 전처리 실행

    Args:
        file_directory (str): 파일이 있는 디렉토리
        hong_file_name (str): HONG 파일명
        output_file_name (str): 출력 파일명
    """
    logger.info("=== HONG Data Excel 파일 - 연간교육계획 시트 전처리 시작 ===")

    # 파일 경로 설정
    file_path = os.path.join(file_directory, hong_file_name)

    # 파일 존재 확인
    if not os.path.exists(file_path):
        logger.error(f"✗ 파일을 찾을 수 없습니다: {file_path}")
        logger.error("파일 경로를 확인해주세요.")
        return False

    logger.info(f"처리할 파일: {file_path}")

    # 1,2단계: Excel 파일 불러오기 및 연간교육계획 시트 찾기
    xl_file = load_hong_excel_file(file_path)
    annual_plan_sheet = find_annual_plan_sheet(xl_file)

    if annual_plan_sheet is None:
        logger.error("✗ 연간교육계획 시트를 찾을 수 없습니다.")
        return False

    # 3단계: 연간교육계획 시트의 컬럼명 추출
    columns = get_annual_plan_columns(file_path, annual_plan_sheet)

    if columns is not None:
        # 4단계: 날짜 컬럼 추가
        df_with_dates = add_date_columns(file_path, annual_plan_sheet)

        if df_with_dates is not None:
            # 5단계: Subsidiary 매핑 수정
            df_with_subsidiary = update_subsidiary_mapping(df_with_dates)

            if df_with_subsidiary is not None:
                # 6단계: Manage Area 필터링
                df_final = filter_hong_manage_area(df_with_subsidiary, file_directory, "index_management_final.csv")

                if df_final is None:
                    logger.error("✗ Manage Area 필터링에 실패했습니다.")
                    return False

                # 최종단계: CSV 파일로 저장
                logger.info("최종단계: 최종 결과를 CSV 파일로 저장합니다...")
                output_path = os.path.join(file_directory, output_file_name)
                df_final.to_csv(output_path, index=False, encoding='utf-8-sig')
                logger.info(f"✓ CSV 파일 저장 완료: {output_path}")
                logger.info(f"✓ 저장된 데이터: {df_final.shape[0]}행, {df_final.shape[1]}열")
                logger.info("=== HONG 연간교육계획 전처리 완료 ===")
                return True
            else:
                logger.error("✗ Subsidiary 매핑 수정에 실패했습니다.")
                return False
        else:
            logger.error("✗ 날짜 컬럼 추가에 실패했습니다.")
            return False
    else:
        logger.error("✗ 연간교육계획 시트 컬럼 추출에 실패했습니다.")
        return False
