#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
메인 실행 스크립트
다양한 전처리 함수들을 호출하여 실행하는 메인 스크립트입니다.
"""

import os
import sys
import pandas as pd
from logger_config import get_default_logger
from excel_preprocess_hr import load_excel_file, find_detail_sheet, get_detail_columns, create_final_company_name, update_region_mp_complete, extract_new_hire_complete, update_branch_mapping, create_new_leader_column, filter_manage_area
from excel_preprocess_lms import get_lms_columns, group_category
from excel_preprocess_hong import run_hong_manager_preprocessing, run_hong_plan_preprocessing
from make_logic import run_make_logic

# ==================== 분석 기준 설정 ====================
# 이 값들만 변경하면 모든 전처리 및 분석이 해당 월 기준으로 수행됩니다
ANALYSIS_YEAR = 2025  # 분석 기준 년도
ANALYSIS_MONTH = 9    # 분석 기준 월 (1~12)
# ========================================================

# 전역 변수 설정
SOURCE_DIRECTORY = f"원본/{ANALYSIS_MONTH}월"  # 원본 파일이 있는 디렉토리
FILE_DIRECTORY = f"data/{ANALYSIS_MONTH}"  # 작업 파일이 있는 디렉토리 (월별로 자동 설정)
INDEX_MANAGEMENT_FILE = "index_management.xlsx"  # Index Management 파일명
HR_FILE_NAME = "hr_index.xlsx"  # HR 처리할 파일명
LMS_FILE_NAME = "lms_learning.xlsx"  # LMS 처리할 파일명
HONG_FILE_NAME = "hong_data.xlsx"  # HONG 처리할 파일명
INDEX_MANAGEMENT_OUTPUT = "index_management_final.csv"  # Index Management 출력 파일명
HR_OUTPUT_FILE_NAME = "hr_index_final.csv"  # HR 최종 출력 파일명
LMS_OUTPUT_FILE_NAME = "lms_learning_final.csv"  # LMS 최종 출력 파일명
HONG_MANAGER_OUTPUT_FILE_NAME = "hong_data_manager_final.csv"  # HONG 법인담당자 출력 파일명
HONG_PLAN_OUTPUT_FILE_NAME = "hong_data_plan_final.csv"  # HONG 연간교육계획 출력 파일명

# 로거 설정
logger = get_default_logger(__name__)

def run_index_management_preprocessing():
    """
    Index Management Excel 파일을 CSV로 변환
    """
    logger.info("=== Index Management Excel 파일 전처리 시작 ===")

    # 파일 경로 (data 디렉토리에서 가져옴)
    source_path = os.path.join(FILE_DIRECTORY, INDEX_MANAGEMENT_FILE)

    # 파일 존재 확인
    if not os.path.exists(source_path):
        logger.error(f"✗ 파일을 찾을 수 없습니다: {source_path}")
        logger.error("파일 경로를 확인해주세요.")
        return False

    logger.info(f"원본 파일: {source_path}")

    try:
        # Excel 파일 읽기
        logger.info("Excel 파일을 읽는 중...")
        df = pd.read_excel(source_path, sheet_name=0)  # 첫 번째 시트
        logger.info(f"✓ Excel 파일 로드 완료: {df.shape[0]}행, {df.shape[1]}열")

        # 컬럼명 확인
        logger.info(f"컬럼 목록: {list(df.columns)}")

        # Y/N 컬럼의 빈 값을 'N'으로 채우기
        logger.info("Y/N 컬럼의 빈 값 처리 중...")
        yn_columns = [
            'Manage Area',
            'New LMS Course',
            'LMS Mission',
            'Annual Plan Setup',
            'JAM Member',
            'Global L&D Council',
            'Infra index response'
        ]

        filled_counts = {}
        for col in yn_columns:
            if col in df.columns:
                # 빈 값 (null, NaN, 빈 문자열) 찾기
                empty_mask = df[col].isna() | (df[col] == '') | (df[col].astype(str).str.strip() == '')
                empty_count = empty_mask.sum()

                if empty_count > 0:
                    filled_counts[col] = empty_count
                    logger.info(f"  - {col}: {empty_count}개 빈 값 → 'N'으로 변경")
                    df.loc[empty_mask, col] = 'N'

        if filled_counts:
            logger.info(f"✓ 총 {sum(filled_counts.values())}개의 빈 값을 'N'으로 채움")
        else:
            logger.info(f"✓ 모든 Y/N 컬럼에 빈 값 없음")

        # CSV 파일로 저장
        output_path = os.path.join(FILE_DIRECTORY, INDEX_MANAGEMENT_OUTPUT)
        logger.info(f"CSV 파일로 저장 중: {output_path}")
        df.to_csv(output_path, index=False, encoding='utf-8-sig')

        logger.info(f"✓ CSV 파일 저장 완료: {output_path}")
        logger.info(f"✓ 저장된 데이터: {df.shape[0]}행, {df.shape[1]}열")
        logger.info("=== Index Management 전처리 완료 ===")
        return True

    except Exception as e:
        logger.error(f"✗ Index Management 전처리 중 오류 발생: {e}")
        return False

def run_prev_hr_preprocessing():
    """
    Prev HR Index Excel 파일을 CSV로 변환 (단순 변환만)
    """
    logger.info("=== Prev HR Index Excel 파일 전처리 시작 ===")

    # 파일 경로
    source_path = os.path.join(FILE_DIRECTORY, "prev_hr_index.xlsx")
    output_path = os.path.join(FILE_DIRECTORY, "prev_hr_index_final.csv")

    # 파일 존재 확인
    if not os.path.exists(source_path):
        logger.error(f"✗ 파일을 찾을 수 없습니다: {source_path}")
        logger.error("파일 경로를 확인해주세요.")
        return False

    logger.info(f"원본 파일: {source_path}")

    try:
        # Excel 파일의 시트 목록 확인
        excel_file = pd.ExcelFile(source_path)
        sheet_names = excel_file.sheet_names
        logger.info(f"사용 가능한 시트: {sheet_names}")

        # 'detail' 시트 찾기 (대소문자 구분 없이)
        detail_sheet = None
        for sheet in sheet_names:
            if sheet.lower() == 'detail':
                detail_sheet = sheet
                break

        if detail_sheet is None:
            logger.error(f"✗ 'Detail' 시트를 찾을 수 없습니다. 사용 가능한 시트: {sheet_names}")
            return False

        # Excel 파일 읽기
        logger.info(f"Excel 파일을 읽는 중... ('{detail_sheet}' 시트)")
        df = pd.read_excel(source_path, sheet_name=detail_sheet)
        logger.info(f"✓ Excel 파일 로드 완료: {df.shape[0]}행, {df.shape[1]}열")

        # 컬럼명 확인
        logger.info(f"컬럼 개수: {len(df.columns)}개")
        logger.info(f"첫 10개 컬럼: {list(df.columns)[:10]}")

        # CSV 파일로 저장
        logger.info(f"CSV 파일로 저장 중: {output_path}")
        df.to_csv(output_path, index=False, encoding='utf-8-sig')

        logger.info(f"✓ CSV 파일 저장 완료: {output_path}")
        logger.info(f"✓ 저장된 데이터: {df.shape[0]}행, {df.shape[1]}열")
        logger.info("=== Prev HR Index 전처리 완료 ===")
        return True

    except Exception as e:
        logger.error(f"✗ Prev HR Index 전처리 중 오류 발생: {e}")
        import traceback
        logger.error(f"상세 오류:\n{traceback.format_exc()}")
        return False

def run_hr_preprocessing():
    """
    HR Index Excel 파일 전처리 실행
    """
    logger.info("=== HR Index Excel 파일 전처리 시작 ===")

    # 전역 변수로 파일 경로 설정
    file_path = os.path.join(FILE_DIRECTORY, HR_FILE_NAME)

    # 파일 존재 확인
    if not os.path.exists(file_path):
        logger.error(f"✗ 파일을 찾을 수 없습니다: {file_path}")
        logger.error("파일 경로를 확인해주세요.")
        return False

    logger.info(f"처리할 파일: {file_path}")

    # 1,2단계: Excel 파일 불러오기 및 DETAIL 시트 찾기 (한 번만 실행)
    xl_file = load_excel_file(file_path)
    detail_sheet = find_detail_sheet(xl_file)

    if detail_sheet is None:
        logger.error("✗ DETAIL 시트를 찾을 수 없습니다.")
        return False

    # detail 시트의 모든 컬럼명 추출
    columns = get_detail_columns(file_path, detail_sheet)

    if columns is not None:
        # 4단계: 최종 법인명 생성
        df_with_final_company = create_final_company_name(file_path, detail_sheet)

        if df_with_final_company is not None:
            # 5단계: Region(MP) 수정
            df_with_region = update_region_mp_complete(df_with_final_company, file_path, detail_sheet)

            if df_with_region is not None:
                # 6단계: New Hire 추출
                df_with_new_hire = extract_new_hire_complete(df_with_region, file_path, detail_sheet, ANALYSIS_YEAR)

                if df_with_new_hire is not None:
                    # 7단계: Branch 매핑 처리
                    df_with_branch = update_branch_mapping(df_with_new_hire, file_path, detail_sheet)

                    if df_with_branch is not None:
                        # 7단계: New Leader 컬럼 생성
                        df_with_new_leader = create_new_leader_column(df_with_branch, FILE_DIRECTORY, ANALYSIS_YEAR)

                        if df_with_new_leader is None:
                            logger.error("✗ New Leader 컬럼 생성에 실패했습니다.")
                            return False

                        # 8단계: Manage Area 필터링
                        df_final = filter_manage_area(df_with_new_leader, FILE_DIRECTORY, INDEX_MANAGEMENT_OUTPUT)

                        if df_final is None:
                            logger.error("✗ Manage Area 필터링에 실패했습니다.")
                            return False

                        # 최종단계: CSV 파일로 저장
                        logger.info("최종단계: 최종 결과를 CSV 파일로 저장합니다...")

                        # 생성된 주요 컬럼 확인
                        logger.info("  생성된 주요 컬럼:")
                        if 'Final Sub.' in df_final.columns:
                            unique_subs = df_final['Final Sub.'].nunique()
                            logger.info(f"    ✓ 'Final Sub.' - {unique_subs}개 고유 법인명")
                        if 'Final Region' in df_final.columns:
                            unique_regions = df_final['Final Region'].nunique()
                            logger.info(f"    ✓ 'Final Region' - {unique_regions}개 고유 지역명")
                        if 'New Hire' in df_final.columns:
                            new_hire_y = len(df_final[df_final['New Hire'] == 'Y'])
                            new_hire_n = len(df_final[df_final['New Hire'] == 'N'])
                            logger.info(f"    ✓ 'New Hire' - Y: {new_hire_y}건, N: {new_hire_n}건")

                        output_path = os.path.join(FILE_DIRECTORY, HR_OUTPUT_FILE_NAME)
                        df_final.to_csv(output_path, index=False, encoding='utf-8-sig')
                        logger.info(f"✓ CSV 파일 저장 완료: {output_path}")
                        logger.info(f"✓ 저장된 데이터: {df_final.shape[0]}행, {df_final.shape[1]}열")
                        logger.info("=== HR 전처리 완료 ===")
                        return True
                    else:
                        logger.error("✗ Branch 매핑 처리에 실패했습니다.")
                        return False
                else:
                    logger.error("✗ New Hire 추출에 실패했습니다.")
                    return False
            else:
                logger.error("✗ Region(MP) 수정에 실패했습니다.")
                return False
        else:
            logger.error("✗ 최종 법인명 생성에 실패했습니다.")
            return False
    else:
        logger.error("✗ DETAIL 시트 컬럼 추출에 실패했습니다.")
        return False

def run_lms_preprocessing():
    """
    LMS Learning Excel 파일 전처리 실행
    """
    logger.info("=== LMS Learning Excel 파일 전처리 시작 ===")

    # 전역 변수로 파일 경로 설정
    file_path = os.path.join(FILE_DIRECTORY, LMS_FILE_NAME)

    # 파일 존재 확인
    if not os.path.exists(file_path):
        logger.error(f"✗ 파일을 찾을 수 없습니다: {file_path}")
        logger.error("파일 경로를 확인해주세요.")
        return False

    logger.info(f"처리할 파일: {file_path}")

    # LMS 컬럼명 추출
    columns = get_lms_columns(file_path)

    if columns is not None:
        # 3단계: Category 그룹핑
        df_with_category = group_category(file_path)

        if df_with_category is not None:
            # 최종단계: CSV 파일로 저장
            logger.info("최종단계: 최종 결과를 CSV 파일로 저장합니다...")
            output_path = os.path.join(FILE_DIRECTORY, LMS_OUTPUT_FILE_NAME)
            df_with_category.to_csv(output_path, index=False, encoding='utf-8-sig')
            logger.info(f"✓ CSV 파일 저장 완료: {output_path}")
            logger.info(f"✓ 저장된 데이터: {df_with_category.shape[0]}행, {df_with_category.shape[1]}열")
            logger.info("=== LMS 전처리 완료 ===")
            return True
        else:
            logger.error("✗ Category 그룹핑에 실패했습니다.")
            return False
    else:
        logger.error("✗ LMS 컬럼 추출에 실패했습니다.")
        return False


def main():
    """
    메인 실행 함수
    """
    logger.info("=== Excel 전처리 시스템 시작 ===")
    logger.info(f"분석 기준 설정: {ANALYSIS_YEAR}년 {ANALYSIS_MONTH}월")
    logger.info(f"데이터 디렉토리: {FILE_DIRECTORY}")
    logger.info("=" * 60)

    try:
        # Index Management 전처리 실행
        index_mgmt_success = run_index_management_preprocessing()

        if index_mgmt_success:
            logger.info("Index Management 전처리가 성공적으로 완료되었습니다.")
        else:
            logger.error("Index Management 전처리 중 오류가 발생했습니다.")
            return False

        # Prev HR Index 전처리 실행
        prev_hr_success = run_prev_hr_preprocessing()

        if prev_hr_success:
            logger.info("Prev HR Index 전처리가 성공적으로 완료되었습니다.")
        else:
            logger.error("Prev HR Index 전처리 중 오류가 발생했습니다.")
            return False

        # HR 전처리 실행
        hr_success = run_hr_preprocessing()

        if hr_success:
            logger.info("HR 전처리가 성공적으로 완료되었습니다.")
        else:
            logger.error("HR 전처리 중 오류가 발생했습니다.")

        # LMS 전처리 실행
        lms_success = run_lms_preprocessing()

        if lms_success:
            logger.info("LMS 전처리가 성공적으로 완료되었습니다.")
        else:
            logger.error("LMS 전처리 중 오류가 발생했습니다.")

        # HONG 법인담당자 전처리 실행
        hong_manager_success = run_hong_manager_preprocessing(
            FILE_DIRECTORY,
            HONG_FILE_NAME,
            HR_OUTPUT_FILE_NAME,
            HONG_MANAGER_OUTPUT_FILE_NAME
        )

        if hong_manager_success:
            logger.info("HONG 법인담당자 전처리가 성공적으로 완료되었습니다.")
        else:
            logger.error("HONG 법인담당자 전처리 중 오류가 발생했습니다.")

        # HONG 연간교육계획 전처리 실행
        hong_plan_success = run_hong_plan_preprocessing(
            FILE_DIRECTORY,
            HONG_FILE_NAME,
            HONG_PLAN_OUTPUT_FILE_NAME
        )

        if hong_plan_success:
            logger.info("HONG 연간교육계획 전처리가 성공적으로 완료되었습니다.")
        else:
            logger.error("HONG 연간교육계획 전처리 중 오류가 발생했습니다.")

        # 로직 생성 실행
        make_logic_success = run_make_logic(FILE_DIRECTORY, ANALYSIS_YEAR, ANALYSIS_MONTH)

        if make_logic_success:
            logger.info("로직 생성이 성공적으로 완료되었습니다.")
        else:
            logger.error("로직 생성 중 오류가 발생했습니다.")

    except Exception as e:
        logger.error(f"메인 실행 중 오류 발생: {e}")
        return False

    logger.info("=== Excel 전처리 시스템 종료 ===")
    return True

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
