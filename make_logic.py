#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
데이터 처리 로직 스크립트
전처리된 파일들을 불러와서 분석 및 처리하는 기능을 제공합니다.
"""

import pandas as pd
import os
import sys
import math
from logger_config import get_default_logger

# 로거 설정
logger = get_default_logger(__name__)

# 전처리된 파일명 설정 (디렉토리는 main.py에서 전달받음)
HR_FINAL_FILE = "hr_index_final.csv"
LMS_FINAL_FILE = "lms_learning_final.csv"
HONG_MANAGER_FINAL_FILE = "hong_data_manager_final.csv"
HONG_PLAN_FINAL_FILE = "hong_data_plan_final.csv"

# 분석 기준 월 설정 (main.py에서 전달받음)
# 이 변수들은 run_make_logic() 함수 내에서 설정됨
ANALYSIS_YEAR = None
ANALYSIS_MONTH = None
ANALYSIS_MONTH_STR = None

def load_processed_files(file_directory):
    """
    전처리된 파일 4개를 불러오는 함수 (1단계)

    Args:
        file_directory (str): 파일이 있는 디렉토리
    """
    try:
        # 1단계: 전처리된 파일들 불러오기
        logger.info("1단계: 전처리된 파일들을 불러옵니다...")

        processed_files = {}

        # HR 파일 불러오기
        hr_path = os.path.join(file_directory, HR_FINAL_FILE)
        if os.path.exists(hr_path):
            logger.info("1.1단계: HR 파일을 불러옵니다...")
            processed_files['hr'] = pd.read_csv(hr_path, encoding='utf-8-sig')
            logger.info(f"✓ HR 파일 불러오기 완료: {processed_files['hr'].shape[0]}행, {processed_files['hr'].shape[1]}열")
        else:
            logger.error(f"✗ HR 파일을 찾을 수 없습니다: {hr_path}")
            return None

        # LMS 파일 불러오기
        lms_path = os.path.join(file_directory, LMS_FINAL_FILE)
        if os.path.exists(lms_path):
            logger.info("1.2단계: LMS 파일을 불러옵니다...")
            processed_files['lms'] = pd.read_csv(lms_path, encoding='utf-8-sig', low_memory=False)
            logger.info(f"✓ LMS 파일 불러오기 완료: {processed_files['lms'].shape[0]}행, {processed_files['lms'].shape[1]}열")
        else:
            logger.error(f"✗ LMS 파일을 찾을 수 없습니다: {lms_path}")
            return None

        # HONG Manager 파일은 make_logic에서 사용하지 않음 (주석 처리)
        # hong_manager_path = os.path.join(file_directory, HONG_MANAGER_FINAL_FILE)
        # if os.path.exists(hong_manager_path):
        #     logger.info("1.3단계: HONG Manager 파일을 불러옵니다...")
        #     processed_files['hong_manager'] = pd.read_csv(hong_manager_path, encoding='utf-8-sig')
        #     logger.info(f"✓ HONG Manager 파일 불러오기 완료: {processed_files['hong_manager'].shape[0]}행, {processed_files['hong_manager'].shape[1]}열")
        # else:
        #     logger.error(f"✗ HONG Manager 파일을 찾을 수 없습니다: {hong_manager_path}")
        #     return None

        # HONG Plan 파일 불러오기
        hong_plan_path = os.path.join(file_directory, HONG_PLAN_FINAL_FILE)
        if os.path.exists(hong_plan_path):
            logger.info("1.3단계: HONG Plan 파일을 불러옵니다...")
            processed_files['hong_plan'] = pd.read_csv(hong_plan_path, encoding='utf-8-sig')
            logger.info(f"✓ HONG Plan 파일 불러오기 완료: {processed_files['hong_plan'].shape[0]}행, {processed_files['hong_plan'].shape[1]}열")
        else:
            logger.error(f"✗ HONG Plan 파일을 찾을 수 없습니다: {hong_plan_path}")
            return None

        logger.info("✓ 모든 전처리된 파일 불러오기 완료")

        # 파일별 컬럼 정보 출력
        logger.info("불러온 파일들의 컬럼 정보:")
        for file_name, df in processed_files.items():
            logger.info(f"  {file_name}: {list(df.columns)}")

        return processed_files

    except Exception as e:
        logger.error(f"✗ 오류 발생: {e}")
        return None

def create_join_table(df_hr, df_lms, file_directory):
    """
    HR과 LMS 테이블을 조인하는 함수 (2단계)

    Args:
        df_hr: HR 데이터프레임
        df_lms: LMS 데이터프레임
        file_directory (str): 파일 저장 디렉토리
    """
    try:
        # 2단계: 조인 테이블 생성
        logger.info("2단계: HR과 LMS 테이블을 조인합니다...")

        # 조인할 컬럼 확인
        hr_join_col = 'Emp. No.'
        lms_join_col = 'Employee Number'

        if hr_join_col not in df_hr.columns:
            logger.error(f"✗ HR 테이블에 '{hr_join_col}' 컬럼이 없습니다.")
            logger.info(f"HR 테이블 컬럼: {list(df_hr.columns)}")
            return None

        if lms_join_col not in df_lms.columns:
            logger.error(f"✗ LMS 테이블에 '{lms_join_col}' 컬럼이 없습니다.")
            logger.info(f"LMS 테이블 컬럼: {list(df_lms.columns)}")
            return None

        logger.info(f"✓ HR 조인 컬럼: {hr_join_col}")
        logger.info(f"✓ LMS 조인 컬럼: {lms_join_col}")

        # 데이터 타입 확인 및 변환
        logger.info("2.1단계: 조인 컬럼 데이터 타입을 확인합니다...")

        # HR Emp. No. 데이터 타입 확인
        hr_emp_no_sample = df_hr[hr_join_col].dropna().head(5)
        logger.info(f"HR Emp. No. 샘플 데이터: {hr_emp_no_sample.tolist()}")

        # LMS Employee Number 데이터 타입 확인
        lms_emp_no_sample = df_lms[lms_join_col].dropna().head(5)
        logger.info(f"LMS Employee Number 샘플 데이터: {lms_emp_no_sample.tolist()}")

        # 사번 중복 및 매칭 문제 분석
        logger.info("2.2단계: 사번 중복 및 매칭 문제를 분석합니다...")

        # HR 사번 중복 확인
        hr_duplicates = df_hr[hr_join_col].duplicated().sum()
        hr_total = len(df_hr[hr_join_col].dropna())
        hr_unique = df_hr[hr_join_col].nunique()
        logger.info(f"HR 사번 통계: 총 {hr_total}개, 고유값 {hr_unique}개, 중복 {hr_duplicates}개")

        if hr_duplicates > 0:
            hr_dup_values = df_hr[df_hr[hr_join_col].duplicated(keep=False)][hr_join_col].value_counts()
            logger.warning(f"HR 중복 사번 (상위 10개):")
            for emp_no, count in hr_dup_values.head(10).items():
                logger.warning(f"  {emp_no}: {count}개")

        # LMS 사번 통계 (중복 계산 제외)
        lms_total = len(df_lms[lms_join_col].dropna())
        lms_unique = df_lms[lms_join_col].nunique()
        logger.info(f"LMS 사번 통계: 총 {lms_total}개, 고유값 {lms_unique}개")

        # 매칭 가능한 사번 확인
        hr_emp_nos = set(df_hr[hr_join_col].dropna().astype(str))
        lms_emp_nos = set(df_lms[lms_join_col].dropna().astype(str))

        common_emp_nos = hr_emp_nos.intersection(lms_emp_nos)
        hr_only = hr_emp_nos - lms_emp_nos
        lms_only = lms_emp_nos - hr_emp_nos

        logger.info(f"매칭 분석:")
        logger.info(f"  공통 사번: {len(common_emp_nos)}개")
        logger.info(f"  HR에만 있는 사번: {len(hr_only)}개")
        logger.info(f"  LMS에만 있는 사번: {len(lms_only)}개")

        if len(hr_only) > 0:
            logger.warning(f"HR에만 있는 사번 샘플 (상위 10개): {list(hr_only)[:10]}")
        if len(lms_only) > 0:
            logger.warning(f"LMS에만 있는 사번 샘플 (상위 10개): {list(lms_only)[:10]}")

        # 매칭률 계산
        hr_match_rate = len(common_emp_nos) / len(hr_emp_nos) * 100 if len(hr_emp_nos) > 0 else 0
        lms_match_rate = len(common_emp_nos) / len(lms_emp_nos) * 100 if len(lms_emp_nos) > 0 else 0

        logger.info(f"매칭률: HR 기준 {hr_match_rate:.2f}%, LMS 기준 {lms_match_rate:.2f}%")

        # 조인 수행 (left join)
        logger.info("2.3단계: Left Join을 수행합니다...")
        logger.info("  조인 방식: Left Join (LMS를 기준으로 HR 데이터를 붙임)")
        logger.info("  설명:")
        logger.info("    - Left 테이블: LMS (모든 행 유지)")
        logger.info("    - Right 테이블: HR (매칭되는 행만 붙음)")
        logger.info(f"    - 조인 키: LMS.'{lms_join_col}' = HR.'{hr_join_col}'")
        logger.info("  결과:")
        logger.info("    - LMS의 모든 교육 이수 기록은 그대로 유지됨")
        logger.info("    - 사번이 매칭되면 HR의 법인정보(Final Sub., Final Region 등)가 추가됨")
        logger.info("    - 사번이 매칭되지 않으면 HR 컬럼들은 null로 채워짐")

        join_table = pd.merge(df_lms, df_hr, left_on=lms_join_col, right_on=hr_join_col, how='left')

        logger.info(f"✓ 조인 완료:")
        logger.info(f"  - 조인 결과: {join_table.shape[0]}행, {join_table.shape[1]}열")
        logger.info(f"  - LMS 원본: {df_lms.shape[0]}행, {df_lms.shape[1]}열 (기준)")
        logger.info(f"  - HR 원본: {df_hr.shape[0]}행, {df_hr.shape[1]}열 (참조)")
        logger.info(f"  - 추가된 컬럼: {join_table.shape[1] - df_lms.shape[1]}개 (HR에서)")

        # 조인 결과 분석
        logger.info("2.4단계: 조인 결과를 분석합니다...")

        # 조인 결과 통계 (LMS 기준)
        join_total = len(join_table)
        join_unique = join_table[lms_join_col].nunique()

        logger.info(f"조인 결과 사번 통계: 총 {join_total}개, 고유값 {join_unique}개")

        # Left Join 결과 분석
        matched_count = join_table[hr_join_col].notna().sum()
        unmatched_count = join_table[hr_join_col].isna().sum()

        logger.info(f"Left Join 결과:")
        logger.info(f"  LMS 데이터: {df_lms.shape[0]}개 (모두 유지)")
        logger.info(f"  HR과 매칭된 LMS: {matched_count}개 ({matched_count/df_lms.shape[0]*100:.2f}%)")
        logger.info(f"  HR과 매칭되지 않은 LMS: {unmatched_count}개 ({unmatched_count/df_lms.shape[0]*100:.2f}%)")
        logger.info(f"  HR에서 제외된 행: {df_hr.shape[0] - matched_count}개")

        # 중복 컬럼 확인 및 처리
        logger.info("2.5단계: 중복 컬럼을 확인합니다...")
        duplicate_cols = [col for col in join_table.columns if col.endswith('_x') or col.endswith('_y')]
        if duplicate_cols:
            logger.info(f"중복 컬럼 발견: {duplicate_cols}")
        else:
            logger.info("✓ 중복 컬럼 없음")

        # Final Sub. 컬럼이 비어있는 레코드 삭제
        logger.info("2.6단계: Final Sub.가 비어있는 레코드를 삭제합니다...")

        # Final Sub. 컬럼 찾기
        final_sub_col = None
        for col in join_table.columns:
            if 'final' in col.lower() and 'sub' in col.lower():
                final_sub_col = col
                break

        if final_sub_col is not None:
            # 삭제 전 상태 확인
            before_count = len(join_table)
            empty_final_sub_count = join_table[final_sub_col].isna().sum()

            logger.info(f"✓ Final Sub. 컬럼 발견: {final_sub_col}")
            logger.info(f"✓ 삭제 전 총 레코드: {before_count}개")
            logger.info(f"✓ Final Sub.가 비어있는 레코드: {empty_final_sub_count}개")

            # Final Sub.가 비어있는 레코드 삭제
            join_table = join_table.dropna(subset=[final_sub_col])

            # 삭제 후 상태 확인
            after_count = len(join_table)
            deleted_count = before_count - after_count
            deletion_rate = (deleted_count / before_count * 100) if before_count > 0 else 0

            logger.info(f"✓ 삭제 후 총 레코드: {after_count}개")
            logger.info(f"✓ 삭제된 레코드: {deleted_count}개 ({deletion_rate:.2f}%)")

            if deleted_count > 0:
                logger.info(f"✓ Final Sub.가 비어있는 레코드 삭제 완료")
                logger.info("")
                logger.info("=" * 80)
                logger.info("【중요】 인사(hr_index)에 없는 교육인원의 교육 관련 자료(lms)는 삭제됨")
                logger.info("=" * 80)
                logger.info("  - 이유: HR 데이터와 매칭되지 않는 LMS 기록은 법인 정보가 없어 분석 불가")
                logger.info("  - 설명:")
                logger.info("    1) Left Join 시 LMS의 사번이 HR에 없으면 Final Sub. = null")
                logger.info("    2) Final Sub.가 null인 레코드는 어느 법인인지 알 수 없음")
                logger.info("    3) 따라서 해당 레코드는 통계 처리에서 제외됨")
                logger.info(f"  - 삭제된 LMS 레코드: {deleted_count}개")
                logger.info(f"  - 전체 LMS 레코드 대비 비율: {deletion_rate:.2f}%")
                logger.info(f"  - 남은 LMS 레코드: {after_count}개 ({100-deletion_rate:.2f}%)")
                logger.info("=" * 80)
                logger.info("")
            else:
                logger.info(f"✓ Final Sub.가 비어있는 레코드가 없습니다 (모든 LMS 기록이 HR과 매칭됨)")
        else:
            logger.warning("✗ Final Sub. 컬럼을 찾을 수 없습니다.")

        # CSV 파일로 저장
        logger.info("2.7단계: 조인 테이블을 CSV 파일로 저장합니다...")
        output_path = os.path.join(file_directory, "join_hr_lms.csv")
        join_table.to_csv(output_path, index=False, encoding='utf-8-sig')

        logger.info(f"✓ 조인 테이블 저장 완료: {output_path}")
        logger.info(f"✓ 저장된 데이터: {join_table.shape[0]}행, {join_table.shape[1]}열")

        logger.info("✓ 조인 테이블 생성 완료")

        return join_table

    except Exception as e:
        logger.error(f"✗ 오류 발생: {e}")
        return None

def calculate_current_education_plans(df_hong_plan):
    """
    현재 시간 기준에 포함되는 교육계획 개수를 계산하는 함수 (3.1단계)
    """
    try:
        # 3.1단계: 이번달 교육계획 개수 계산
        logger.info("3.1단계: 이번달 교육계획 개수를 계산합니다...")
        logger.info("  - 사용 파일: hong_data_plan_final.csv (HONG 연간교육계획)")
        logger.info("  - 목적: 법인별로 현재 진행 중인 교육 과정 수를 파악 (계획 대비 실행률 계산용)")

        # 분석 기준 월 가져오기
        analysis_time = pd.Timestamp(f"{ANALYSIS_YEAR}-{ANALYSIS_MONTH:02d}-01")
        logger.info(f"  - 분석 기준 시점: {ANALYSIS_YEAR}년 {ANALYSIS_MONTH}월 1일 ({analysis_time})")

        # 필요한 컬럼 확인
        required_columns = ['Subsidiary', 'Start Date', 'End Date']
        missing_columns = [col for col in required_columns if col not in df_hong_plan.columns]

        if missing_columns:
            logger.error(f"✗ 필요한 컬럼이 없습니다: {missing_columns}")
            return None

        logger.info(f"  - 필요한 컬럼 확인 완료: {required_columns}")

        # Start Date, End Date를 datetime으로 변환
        logger.info("  - Start Date, End Date를 datetime으로 변환 중...")
        logger.info("    (형식: YYYYMMDD → datetime)")
        df_hong_plan['Start Date'] = pd.to_datetime(df_hong_plan['Start Date'], format='%Y%m%d', errors='coerce')
        df_hong_plan['End Date'] = pd.to_datetime(df_hong_plan['End Date'], format='%Y%m%d', errors='coerce')

        # 분석 기준 월이 Start Date와 End Date 사이에 있는 조건
        logger.info("  - 이번달 교육계획 필터링 조건:")
        logger.info(f"    조건: Start Date <= 분석기준일({analysis_time}) <= End Date")
        logger.info("    의미: 분석 기준 시점(8월 1일)에 진행 중인 과정만 선택")
        logger.info("    예시:")
        logger.info("      ✓ 포함: Start Date = 2025-07-01, End Date = 2025-09-30 (8월 포함)")
        logger.info("      ✓ 포함: Start Date = 2025-08-01, End Date = 2025-08-31 (8월 전체)")
        logger.info("      ✓ 포함: Start Date = 2025-01-01, End Date = 2025-12-31 (연중 과정)")
        logger.info("      ✗ 제외: Start Date = 2025-09-01, End Date = 2025-09-30 (9월 과정)")
        logger.info("      ✗ 제외: Start Date = 2025-06-01, End Date = 2025-07-31 (7월까지)")

        current_condition = (
            (df_hong_plan['Start Date'] <= analysis_time) &
            (df_hong_plan['End Date'] >= analysis_time)
        )

        # 현재 시간 기준 교육계획 필터링
        current_plans = df_hong_plan[current_condition].copy()

        total_plans = len(df_hong_plan)
        current_plans_count = len(current_plans)
        filtered_out = total_plans - current_plans_count

        logger.info(f"  - 필터링 결과:")
        logger.info(f"    전체 교육계획: {total_plans}개")
        logger.info(f"    이번달(8월) 진행 중: {current_plans_count}개 ({current_plans_count/total_plans*100:.1f}%)")
        logger.info(f"    제외됨 (기간 불일치): {filtered_out}개 ({filtered_out/total_plans*100:.1f}%)")

        # 법인별 개수 계산
        subsidiary_counts = current_plans['Subsidiary'].value_counts()

        logger.info("법인별 이번달 진행 중인 교육계획 개수:")
        for subsidiary, count in subsidiary_counts.items():
            logger.info(f"  {subsidiary}: {count}개")

        logger.info("✓ 이번달 교육계획 개수 계산 완료")

        return {
            'current_plans': current_plans,
            'subsidiary_counts': subsidiary_counts,
            'total_current_plans': len(current_plans)
        }

    except Exception as e:
        logger.error(f"✗ 오류 발생: {e}")
        return None

def calculate_completed_courses_by_subsidiary(join_table):
    """
    Final Sub.별로 완료된 과정 개수를 계산하는 함수 (3.2단계)
    """
    try:
        # 3.2단계: Final Sub.별 완료된 과정 개수 계산
        logger.info("3.2단계: Final Sub.별 완료된 과정 개수를 계산합니다...")
        logger.info("  목적: 법인별로 실제 완료된 고유 과정 수를 파악 (LMS 과정 등록률 계산용)")

        # 필요한 컬럼 확인
        required_columns = ['Final Sub.', 'Completion status', 'Course name']
        missing_columns = [col for col in required_columns if col not in join_table.columns]

        if missing_columns:
            logger.error(f"✗ 필요한 컬럼이 없습니다: {missing_columns}")
            return None

        logger.info(f"  - 필요한 컬럼 확인 완료: {required_columns}")

        # Completion status가 '-C'로 끝나고, Completion Date가 분석 기준 월인 데이터 필터링
        logger.info("3.2.1단계: 완료된 과정을 필터링합니다...")
        logger.info("  - 필터링 조건 (2가지 모두 만족해야 함):")
        logger.info(f"    조건1) Completion status가 '-C'로 끝나야 함 (완료 상태)")
        logger.info(f"    조건2) Completion Date의 월(MM)이 {ANALYSIS_MONTH}월이어야 함")
        logger.info("  - 완료 상태 판단 기준:")
        logger.info("    Completion status가 '-C'로 끝나면 → 완료된 과정")
        logger.info("    예: 'Online-C', 'Offline-C' 등")
        logger.info("    제외: '수강중', '미수강' 등")

        # Completion status 컬럼의 데이터 타입 확인
        completion_status_sample = join_table['Completion status'].dropna().head(10)
        logger.info(f"  - Completion status 샘플 (처음 10개): {completion_status_sample.tolist()}")

        # 고유 상태값 확인
        unique_statuses = join_table['Completion status'].unique()
        logger.info(f"  - Completion status 고유값 ({len(unique_statuses)}개): {list(unique_statuses)}")

        # Completion Date가 있는지 확인
        if 'Completion Date' not in join_table.columns:
            logger.error("✗ Completion Date 컬럼을 찾을 수 없습니다.")
            return None

        # Completion Date를 datetime으로 변환
        logger.info("  - Completion Date를 datetime으로 변환 중...")
        join_table_copy = join_table.copy()
        join_table_copy['Completion Date'] = join_table_copy['Completion Date'].astype(str).str.replace('.0', '')
        join_table_copy['Completion Date'] = pd.to_datetime(join_table_copy['Completion Date'], format='%Y%m%d', errors='coerce')

        # '-C'로 끝나는 조건
        completed_condition = join_table_copy['Completion status'].str.endswith('-C', na=False)

        # 해당 월 조건 (Completion Date의 월이 ANALYSIS_MONTH와 같음)
        month_condition = join_table_copy['Completion Date'].dt.month == ANALYSIS_MONTH

        # 두 조건을 모두 만족하는 데이터 필터링
        completed_courses = join_table_copy[completed_condition & month_condition].copy()

        # 각 조건별 통계
        total_records = len(join_table_copy)
        completed_all = len(join_table_copy[completed_condition])
        completed_this_month = len(completed_courses)
        not_this_month = completed_all - completed_this_month

        logger.info(f"  - 필터링 결과:")
        logger.info(f"    전체 LMS 레코드: {total_records}개")
        logger.info(f"    완료 상태 레코드 (모든 월): {completed_all}개 ({completed_all/total_records*100:.1f}%)")
        logger.info(f"    완료 상태 + {ANALYSIS_MONTH}월 완료: {completed_this_month}개 ({completed_this_month/total_records*100:.1f}%)")
        logger.info(f"    완료 상태이지만 다른 월: {not_this_month}개 ({not_this_month/completed_all*100:.1f}% of 완료)")

        if len(completed_courses) == 0:
            logger.warning("완료된 과정이 없습니다.")
            return None

        # Final Sub.별 그룹핑 및 Course name 중복 제거 후 개수 계산
        logger.info("3.2.2단계: Final Sub.별로 Course name 중복 제거 후 개수를 계산합니다...")

        # Final Sub.별로 그룹핑하여 Course name의 고유 개수 계산
        subsidiary_course_counts = completed_courses.groupby('Final Sub.')['Course name'].nunique().sort_values(ascending=False)

        logger.info("Final Sub.별 완료된 과정 개수:")
        for subsidiary, count in subsidiary_course_counts.items():
            logger.info(f"  {subsidiary}: {count}개")

        # 전체 통계
        total_completed_courses = subsidiary_course_counts.sum()
        logger.info(f"✓ 전체 완료된 과정 개수: {total_completed_courses}개")
        logger.info("✓ 3.2단계 완료: Final Sub.별 완료된 과정 개수 계산 완료")

        return {
            'subsidiary_course_counts': subsidiary_course_counts,
            'total_completed_courses': total_completed_courses,
            'completed_courses_data': completed_courses
        }

    except Exception as e:
        logger.error(f"✗ 오류 발생: {e}")
        return None

def calculate_completion_rate(current_education_result, completed_courses_result):
    """
    완료율을 계산하는 함수 (3.3단계)
    """
    try:
        # 3.3단계: 완료율 계산
        logger.info("3.3단계: 완료율을 계산합니다...")

        # 3.1 결과에서 법인별 교육계획 개수 추출
        planned_counts = current_education_result['subsidiary_counts']
        logger.info("3.3.1단계: 3.1단계 결과에서 계획 개수를 추출합니다...")
        logger.info(f"계획된 교육계획 개수: {dict(planned_counts)}")

        # 3.2 결과에서 법인별 완료된 과정 개수 추출
        completed_counts = completed_courses_result['subsidiary_course_counts']
        logger.info("3.3.2단계: 3.2단계 결과에서 완료 개수를 추출합니다...")
        logger.info(f"완료된 과정 개수: {dict(completed_counts)}")

        # 법인명 공백 제거 (대문자는 이미 통일되어 있음)
        logger.info("3.3.3단계: 법인명 공백을 제거합니다...")
        logger.info("  - 배경: HR과 HONG에서 모든 법인명이 이미 대문자로 통일되어 있음")
        logger.info("  - 목적: 앞뒤 공백 차이로 인해 같은 법인이 다르게 인식되는 것을 방지")
        logger.info("  - 방법: 앞뒤 공백만 제거 (.strip())")

        # 모든 법인명 수집
        unique_original_names = set(planned_counts.index) | set(completed_counts.index)
        logger.info(f"  - 공백 제거 전 고유 법인명: {len(unique_original_names)}개")

        # 매핑 테이블 생성
        normalization_map = {}  # 공백 제거된 이름 → 대표 원본 이름
        normalized_planned = {}
        normalized_completed = {}

        # 계획 데이터 공백 제거
        logger.info("  - 계획 데이터 공백 제거 중...")
        for subsidiary in planned_counts.index:
            normalized_name = str(subsidiary).strip()  # 공백만 제거

            # 매핑 테이블 업데이트 (첫 번째 원본 이름을 대표로 사용)
            if normalized_name not in normalization_map:
                normalization_map[normalized_name] = subsidiary

            if normalized_name in normalized_planned:
                normalized_planned[normalized_name] += planned_counts[subsidiary]
            else:
                normalized_planned[normalized_name] = planned_counts[subsidiary]

        # 완료 데이터 공백 제거
        logger.info("  - 완료 데이터 공백 제거 중...")
        for subsidiary in completed_counts.index:
            normalized_name = str(subsidiary).strip()  # 공백만 제거

            # 매핑 테이블 업데이트
            if normalized_name not in normalization_map:
                normalization_map[normalized_name] = subsidiary

            if normalized_name in normalized_completed:
                normalized_completed[normalized_name] += completed_counts[subsidiary]
            else:
                normalized_completed[normalized_name] = completed_counts[subsidiary]

        # 공백 제거된 법인 목록
        all_subsidiaries = set(normalized_planned.keys()) | set(normalized_completed.keys())

        logger.info(f"  - 공백 제거 결과:")
        logger.info(f"    처리 전 고유 법인 수: {len(unique_original_names)}개")
        logger.info(f"    처리 후 고유 법인 수: {len(all_subsidiaries)}개")
        logger.info(f"    통합된 법인 수: {len(unique_original_names) - len(all_subsidiaries)}개 (공백으로 인한 중복)")

        # 통합된 법인들 확인 (공백으로 인해 합쳐진 법인들)
        if len(unique_original_names) > len(all_subsidiaries):
            logger.info("  - 공백 제거로 통합된 법인명:")
            # 같은 공백 제거 이름으로 통합된 법인들 찾기
            for normalized_name in normalization_map.keys():
                original_names = [name for name in unique_original_names if str(name).strip() == normalized_name]
                if len(original_names) > 1:
                    logger.info(f"    {original_names} → '{normalization_map[normalized_name]}'")

        logger.info(f"  - 매핑 테이블 크기: {len(normalization_map)}개")
        logger.info("✓ 3.3.3단계 완료: 법인명 공백 제거 완료")

        # 법인별 완료율 계산
        logger.info("3.3.4단계: 법인별 완료율을 계산합니다...")

        completion_rates = {}
        matched_subsidiaries = []
        planned_only = []
        completed_only = []

        for subsidiary in all_subsidiaries:
            planned = normalized_planned.get(subsidiary, 0)
            completed = normalized_completed.get(subsidiary, 0)

            if planned > 0 and completed > 0:
                # 매칭되는 경우: 완료율 계산
                rate = (completed / planned) * 100
                # 소수 2번째 자리에서 올림
                import math
                rounded_rate = math.ceil(rate * 100) / 100
                completion_rates[subsidiary] = rounded_rate
                matched_subsidiaries.append(subsidiary)

                logger.info(f"  {subsidiary}: {completed}/{planned} = {rounded_rate:.2f}%")

            elif planned > 0 and completed == 0:
                # 계획만 있는 경우 → 0%로 처리
                completion_rates[subsidiary] = 0.00
                planned_only.append(subsidiary)
                logger.info(f"  {subsidiary}: {completed}/{planned} = 0.00% (계획만 있음 → 0%로 처리)")

            elif planned == 0 and completed > 0:
                # 완료만 있는 경우 → 0%로 처리
                completion_rates[subsidiary] = 0.00
                completed_only.append(subsidiary)
                logger.info(f"  {subsidiary}: {completed}/0 = 0.00% (완료만 있음 → 0%로 처리)")

        # 결과 요약
        logger.info("3.3.4단계: 결과를 요약합니다...")

        logger.info("매칭 결과:")
        logger.info(f"  매칭되는 법인: {len(matched_subsidiaries)}개")
        logger.info(f"  계획만 있는 법인: {len(planned_only)}개")
        logger.info(f"  완료만 있는 법인: {len(completed_only)}개")

        if planned_only:
            logger.warning(f"계획만 있는 법인들: {planned_only}")
        if completed_only:
            logger.warning(f"완료만 있는 법인들: {completed_only}")

        # 완료율 통계 (모든 법인이 0% 이상의 유효한 값을 가짐)
        valid_rates = list(completion_rates.values())
        if valid_rates:
            avg_rate = sum(valid_rates) / len(valid_rates)
            max_rate = max(valid_rates)
            min_rate = min(valid_rates)

            logger.info("완료율 통계:")
            logger.info(f"  평균 완료율: {avg_rate:.2f}%")
            logger.info(f"  최고 완료율: {max_rate:.2f}%")
            logger.info(f"  최저 완료율: {min_rate:.2f}%")

        # 상위/하위 완료율 법인
        sorted_rates = sorted([(k, v) for k, v in completion_rates.items() if v is not None],
                             key=lambda x: x[1], reverse=True)

        if sorted_rates:
            logger.info("완료율 순위 (상위 5개):")
            for i, (subsidiary, rate) in enumerate(sorted_rates[:5], 1):
                logger.info(f"  {i}위: {subsidiary} ({rate:.2f}%)")

            logger.info("완료율 순위 (하위 5개):")
            for i, (subsidiary, rate) in enumerate(sorted_rates[-5:], len(sorted_rates)-4):
                logger.info(f"  {i}위: {subsidiary} ({rate:.2f}%)")

        # 최종 리스트 생성 (매칭되는 법인 + 계획만 있는 법인)
        logger.info("3.3.5단계: 최종 완료율 리스트를 생성합니다...")

        final_list = []
        processed_subsidiaries = set()  # 이미 처리된 법인들 (정규화된 이름 기준)

        # 매칭되는 법인 추가 (완료율 기준 정렬)
        for subsidiary, rate in sorted_rates:
            if subsidiary not in processed_subsidiaries:
                planned = normalized_planned.get(subsidiary, 0)
                completed = normalized_completed.get(subsidiary, 0)
                # 정규화된 이름을 원본 이름으로 변환
                original_name = normalization_map.get(subsidiary, subsidiary)
                final_list.append({
                    'subsidiary': original_name,
                    'completion_rate': rate,
                    'planned': planned,
                    'completed': completed,
                    'status': '매칭'
                })
                processed_subsidiaries.add(subsidiary)

        # 계획만 있는 법인 추가 (0.00%로 처리) - 이미 처리된 법인은 제외
        for subsidiary in planned_only:
            if subsidiary not in processed_subsidiaries:
                planned = normalized_planned.get(subsidiary, 0)
                # 정규화된 이름을 원본 이름으로 변환
                original_name = normalization_map.get(subsidiary, subsidiary)
                final_list.append({
                    'subsidiary': original_name,
                    'completion_rate': 0.00,
                    'planned': planned,
                    'completed': 0,
                    'status': '계획만 있음'
                })
                processed_subsidiaries.add(subsidiary)

        # 중복 체크
        logger.info("3.3.6단계: 법인 중복을 체크합니다...")
        subsidiary_names = [item['subsidiary'] for item in final_list]
        unique_names = set(subsidiary_names)

        if len(subsidiary_names) != len(unique_names):
            logger.warning("✗ 중복된 법인이 발견되었습니다!")

            # 중복된 법인 찾기
            from collections import Counter
            name_counts = Counter(subsidiary_names)
            duplicates = {name: count for name, count in name_counts.items() if count > 1}

            logger.warning(f"중복된 법인들:")
            for name, count in duplicates.items():
                logger.warning(f"  {name}: {count}번 중복")
        else:
            logger.info("✓ 중복된 법인이 없습니다.")

        # 최종 리스트 출력
        logger.info("최종 완료율 리스트 (높은 순서):")
        for i, item in enumerate(final_list, 1):
            if item['status'] == '매칭':
                logger.info(f"  {i}위: {item['subsidiary']} - {item['completion_rate']:.2f}% ({item['completed']}/{item['planned']})")
            else:
                logger.info(f"  {i}위: {item['subsidiary']} - 0.00% ({item['completed']}/{item['planned']}) [계획만 있음]")

        logger.info("✓ 완료율 계산 완료")

        return {
            'completion_rates': completion_rates,
            'matched_subsidiaries': matched_subsidiaries,
            'planned_only': planned_only,
            'completed_only': completed_only,
            'valid_rates': valid_rates,
            'final_list': final_list,
            'summary_stats': {
                'avg_rate': avg_rate if valid_rates else 0,
                'max_rate': max_rate if valid_rates else 0,
                'min_rate': min_rate if valid_rates else 0
            }
        }

    except Exception as e:
        logger.error(f"✗ 오류 발생: {e}")
        return None

def calculate_monthly_learning_hours(df_hong_plan):
    """
    월별 법인별 Learning Hrs. 계획 시간을 계산하는 함수 (4단계)
    """
    try:
        # 4단계: 월별 법인별 Learning Hrs. 계획 시간 계산
        logger.info("=" * 80)
        logger.info("4단계: 월별 법인별 Learning Hrs. 계획 시간을 계산합니다...")
        logger.info("=" * 80)
        logger.info("사용 파일: hong_data_plan_final.csv")
        logger.info(f"전체 레코드 수: {len(df_hong_plan)}개")
        logger.info("-" * 80)

        # 필요한 컬럼 확인
        required_columns = ['Start Date', 'End Date', 'Subsidiary', 'Learning Hrs.', 'Month_end']
        missing_columns = [col for col in required_columns if col not in df_hong_plan.columns]

        if missing_columns:
            logger.error(f"✗ 필요한 컬럼이 없습니다: {missing_columns}")
            return None

        logger.info(f"✓ 필요한 컬럼 확인 완료: {required_columns}")

        # 데이터 타입 확인 및 변환
        logger.info("4.1.1단계: 날짜 및 시간 데이터를 처리합니다...")
        logger.info("  - 목적: 월별 계획 시간 계산을 위한 필수 컬럼 데이터 변환 및 검증")
        logger.info("  - 처리 내용:")
        logger.info("    1) Start Date: YYYYMMDD 형식 → datetime 변환")
        logger.info("    2) End Date: YYYYMMDD 형식 → datetime 변환")
        logger.info("    3) Learning Hrs.: 문자열 → 숫자 변환")

        # Start Date, End Date를 datetime으로 변환
        df_hong_plan['Start Date'] = pd.to_datetime(df_hong_plan['Start Date'], format='%Y%m%d', errors='coerce')
        df_hong_plan['End Date'] = pd.to_datetime(df_hong_plan['End Date'], format='%Y%m%d', errors='coerce')

        # Learning Hrs.를 숫자로 변환
        df_hong_plan['Learning Hrs.'] = pd.to_numeric(df_hong_plan['Learning Hrs.'], errors='coerce')

        # 샘플 데이터 확인
        logger.info("  - 변환 결과 샘플:")
        logger.info("    Start Date 샘플: " + str(df_hong_plan['Start Date'].dropna().head(3).tolist()))
        logger.info("    End Date 샘플: " + str(df_hong_plan['End Date'].dropna().head(3).tolist()))
        logger.info("    Learning Hrs. 샘플: " + str(df_hong_plan['Learning Hrs.'].dropna().head(3).tolist()))

        # 유효한 데이터만 필터링
        logger.info("  - 유효성 검사:")
        logger.info("    조건: Start Date, End Date, Learning Hrs. 모두 null이 아닌 레코드만 선택")
        logger.info("    이유: 월별 계획 시간 계산을 위해서는 시작일, 종료일, 학습시간 정보가 모두 필요")
        logger.info("    제외 대상: 위 3개 컬럼 중 하나라도 null이거나 변환 실패한 레코드")

        valid_data = df_hong_plan.dropna(subset=['Start Date', 'End Date', 'Learning Hrs.'])

        invalid_count = len(df_hong_plan) - len(valid_data)
        logger.info(f"  - 검사 결과:")
        logger.info(f"    유효한 데이터: {len(valid_data)}개")
        logger.info(f"    제외된 데이터: {invalid_count}개")
        logger.info(f"    전체 데이터: {len(df_hong_plan)}개")
        logger.info(f"    유효 비율: {len(valid_data)/len(df_hong_plan)*100:.1f}%")

        if len(valid_data) == 0:
            logger.warning("유효한 데이터가 없습니다.")
            return None

        # 분석 기준 월 사용
        logger.info(f"✓ 분석 기준: {ANALYSIS_YEAR}년 {ANALYSIS_MONTH}월")

        # 월별 데이터 생성 (새로운 방식: Month_end 기준)
        logger.info(f"4.1.2단계: {ANALYSIS_MONTH}월까지의 계획 시간을 계산합니다...")
        logger.info("  - Month_end란?")
        logger.info("    의미: 각 교육 계획이 완료되어야 하는 목표 월")
        logger.info("    출처: hong_data_plan_final.csv의 Month_end 컬럼")
        logger.info("    예시: Month_end = 8 → 8월까지 완료해야 하는 교육 계획")
        logger.info("  - 계산 로직:")
        logger.info(f"    Month_end가 {ANALYSIS_MONTH} 이하인 교육 계획만 현재 월의 '계획'으로 간주")
        logger.info(f"    예: 분석 기준월이 8월이면 Month_end가 1~8월인 계획만 포함")
        logger.info(f"    이유: 9월 이후 완료 예정 계획은 8월 기준으로 아직 계획 시점이 아님")

        # Month_end 컬럼 확인
        if 'Month_end' not in df_hong_plan.columns:
            logger.error("✗ Month_end 컬럼이 없습니다.")
            return None

        # Month_end를 숫자로 변환
        df_hong_plan['Month_end'] = pd.to_numeric(df_hong_plan['Month_end'], errors='coerce')

        # 유효한 Month_end 데이터만 필터링
        valid_month_data = df_hong_plan.dropna(subset=['Month_end', 'Learning Hrs.', 'Subsidiary'])
        logger.info(f"  - 유효한 Month_end 데이터: {len(valid_month_data)}개")

        # Month_end 분포 확인 (1월부터 12월까지 전체)
        month_end_dist = valid_month_data['Month_end'].value_counts().sort_index()
        logger.info("  - Month_end 분포 (1월~12월):")
        for month in range(1, 13):
            count = month_end_dist.get(float(month), 0)
            if count > 0:
                status = f"✓ 포함" if month <= ANALYSIS_MONTH else "✗ 제외"
                logger.info(f"    {month}월: {int(count)}개 교육 계획 ({status})")
            else:
                logger.info(f"    {month}월: 0개 교육 계획")

        monthly_data = []

        for _, row in valid_month_data.iterrows():
            month_end = int(row['Month_end'])
            subsidiary = row['Subsidiary']
            learning_hrs = row['Learning Hrs.']

            # ANALYSIS_MONTH 이하인 월만 포함
            if month_end <= ANALYSIS_MONTH:
                monthly_data.append({
                    'month': f"{ANALYSIS_YEAR}-{month_end:02d}",
                    'subsidiary': subsidiary,
                    'learning_hrs': learning_hrs,
                    'month_end': month_end
                })

        # 월별 데이터를 DataFrame으로 변환
        monthly_df = pd.DataFrame(monthly_data)

        logger.info(f"✓ 월별 데이터 생성 완료: {len(monthly_df)}개 레코드")

        # 샘플 데이터 출력
        if len(monthly_df) > 0:
            sample_data = monthly_df.head(3)
            for _, row in sample_data.iterrows():
                logger.info(f"  샘플: {row['subsidiary']} - {row['month']}월, {row['learning_hrs']}시간")

        # 법인명 공백 제거
        logger.info("4.1.3단계: 법인명 공백을 제거합니다...")
        logger.info("  - 방법: 앞뒤 공백만 제거 (대문자는 이미 통일되어 있음)")

        # 매핑 테이블 생성
        normalization_map = {}  # 공백 제거된 이름 → 대표 원본 이름
        normalized_monthly_data = []

        for _, row in monthly_df.iterrows():
            subsidiary = row['subsidiary']
            normalized_name = str(subsidiary).strip()  # 공백만 제거

            # 매핑 테이블 업데이트 (첫 번째 원본 이름을 대표로 사용)
            if normalized_name not in normalization_map:
                normalization_map[normalized_name] = subsidiary

            normalized_monthly_data.append({
                'month': row['month'],
                'subsidiary': normalized_name,
                'learning_hrs': row['learning_hrs'],
                'month_end': row['month_end']
            })

        # 정규화된 데이터를 DataFrame으로 변환
        normalized_monthly_df = pd.DataFrame(normalized_monthly_data)

        # 월별 법인별 Learning Hrs. 합계 계산
        logger.info("4.1.4단계: 월별 법인별 Learning Hrs. 합계를 계산합니다...")

        monthly_subsidiary_summary = normalized_monthly_df.groupby(['month', 'subsidiary'])['learning_hrs'].sum().reset_index()

        # 피벗 테이블 생성 (월별로 법인들의 시간 합계)
        pivot_table = monthly_subsidiary_summary.pivot_table(
            index='month',
            columns='subsidiary',
            values='learning_hrs',
            fill_value=0
        )

        logger.info(f"✓ 월별 법인별 계획 시간 계산 완료")
        logger.info(f"✓ 분석 기간: {len(pivot_table)}개월")
        logger.info(f"✓ 분석 법인: {len(pivot_table.columns)}개")
        logger.info("  - 분석 법인이란?")
        logger.info(f"    hong_data_plan_final.csv에서 Month_end가 1~{ANALYSIS_MONTH}월인 레코드 기준")
        logger.info("    법인별로 Learning Hrs. 합산 → 합계가 0보다 큰 법인만 포함")
        logger.info(f"    전체 법인(99개) 중 1~{ANALYSIS_MONTH}월 누적 교육 계획이 있는 법인: {len(pivot_table.columns)}개")
        logger.info(f"    나머지 {99 - len(pivot_table.columns)}개 법인: 1~{ANALYSIS_MONTH}월 누적 교육 계획이 0 또는 데이터 없음")

        # 월별 총 계획 시간 계산
        monthly_totals = pivot_table.sum(axis=1)

        # 법인별 총 계획 시간 계산
        subsidiary_totals = pivot_table.sum(axis=0).sort_values(ascending=False)

        # 결과 출력
        logger.info("4.1.5단계: 결과를 요약합니다...")

        # 월별 총 계획 시간
        logger.info("월별 총 계획 시간:")
        for month, total_hrs in monthly_totals.items():
            logger.info(f"  {month}: {total_hrs:.2f}시간")

        # 일년 총 계획 시간
        yearly_total = monthly_totals.sum()
        logger.info(f"일년 총 계획 시간: {yearly_total:.2f}시간")

        # 법인별 8월 누적 계획 수강 시간 (새로운 방식)
        logger.info(f"법인별 {ANALYSIS_MONTH}월 누적 계획 수강 시간:")

        # 모든 월의 데이터를 합산 (누적)
        cumulative_totals = pivot_table.sum(axis=0).sort_values(ascending=False)

        # 법인별 누적 계획 수강 시간 출력 (정규화된 이름을 원본 이름으로 변환)
        for subsidiary, total_hrs in cumulative_totals.items():
            original_name = normalization_map.get(subsidiary, subsidiary)
            logger.info(f"  {original_name}: {round(total_hrs, 2)}시간")


        logger.info("✓ 월별 법인별 Learning Hrs. 계획 시간 계산 완료")

        return {
            'monthly_subsidiary_data': monthly_subsidiary_summary,
            'pivot_table': pivot_table,
            'monthly_totals': monthly_totals,
            'subsidiary_totals': subsidiary_totals,
            'yearly_total': yearly_total,
            'total_months': len(pivot_table),
            'total_subsidiaries': len(pivot_table.columns),
            'normalization_map': normalization_map
        }

    except Exception as e:
        logger.error(f"✗ 오류 발생: {e}")
        return None

def calculate_monthly_actual_hours(join_table):
    """
    법인별 월별 실제 수강 시간을 계산하는 함수 (4.2단계)
    """
    try:
        # 4.2단계: 법인별 월별 실제 수강 시간 계산
        logger.info("=" * 80)
        logger.info("4.2단계: 법인별 월별 실제 수강 시간을 계산합니다...")
        logger.info("=" * 80)
        logger.info("사용 파일: join_hr_lms.csv")
        logger.info(f"전체 레코드 수: {len(join_table)}개")
        logger.info("-" * 80)

        # 필요한 컬럼 확인
        required_columns = ['Completion status', 'Completion Date', 'Final Sub.', 'Education Hours', 'Staff/Operator']
        missing_columns = [col for col in required_columns if col not in join_table.columns]

        if missing_columns:
            logger.error(f"✗ 필요한 컬럼이 없습니다: {missing_columns}")
            return None

        logger.info(f"✓ 필요한 컬럼 확인 완료: {required_columns}")

        # Completion status가 '-C'로 끝나고 Staff/Operator가 'Staff'인 데이터 필터링
        logger.info("4.2.1단계: 완료된 과정을 필터링합니다...")

        # Completion status 샘플 확인
        logger.info("Completion status 샘플: " + str(join_table['Completion status'].value_counts().head(10).to_dict()))

        # Staff/Operator 샘플 확인
        logger.info("Staff/Operator 샘플: " + str(join_table['Staff/Operator'].value_counts().head(10).to_dict()))

        completed_condition = join_table['Completion status'].str.endswith('-C', na=False)
        staff_condition = join_table['Staff/Operator'] == 'Staff'

        # 두 조건을 모두 만족하는 데이터 필터링
        completed_courses = join_table[completed_condition & staff_condition].copy()

        logger.info(f"✓ 완료된 과정 데이터 (Staff만): {len(completed_courses)}개")

        if len(completed_courses) == 0:
            logger.warning("완료된 과정이 없습니다.")
            return None

        # Completion Date 처리 및 기준월 필터링
        logger.info("4.2.2단계: 기준월까지의 완료 데이터를 필터링합니다...")
        logger.info(f"  - 필터링 조건:")
        logger.info(f"    1) Completion status: '-C'로 끝남 (이미 필터링됨)")
        logger.info(f"    2) Staff/Operator: 'Staff'만 (이미 필터링됨)")
        logger.info(f"    3) Completion Date: {ANALYSIS_YEAR}년 1월 ~ {ANALYSIS_MONTH}월")
        logger.info(f"    4) Education Hours: 숫자로 변환 가능")
        logger.info(f"    5) Final Sub.: null이 아님")
        logger.info(f"  - 목적: {ANALYSIS_YEAR}년 {ANALYSIS_MONTH}월까지 실제로 교육을 완료한 누적 시간 계산")

        # Completion Date를 datetime으로 변환
        completed_courses['Completion Date'] = completed_courses['Completion Date'].astype(str).str.replace('.0', '')
        completed_courses['Completion Date'] = pd.to_datetime(completed_courses['Completion Date'], format='%Y%m%d', errors='coerce')

        # Education Hours를 숫자로 변환
        completed_courses['Education Hours'] = pd.to_numeric(completed_courses['Education Hours'], errors='coerce')

        # 유효한 데이터만 필터링 (null 제거)
        valid_data = completed_courses.dropna(subset=['Completion Date', 'Education Hours', 'Final Sub.']).copy()

        logger.info(f"  - Null 제거 결과:")
        logger.info(f"    입력: {len(completed_courses)}개")
        logger.info(f"    출력: {len(valid_data)}개")
        logger.info(f"    제외: {len(completed_courses) - len(valid_data)}개")

        if len(valid_data) == 0:
            logger.warning("유효한 완료 데이터가 없습니다.")
            return None

        # 기준월 필터링: ANALYSIS_YEAR년 1월 ~ ANALYSIS_MONTH월
        logger.info(f"  - 기준월 필터링: {ANALYSIS_YEAR}년 1월 ~ {ANALYSIS_MONTH}월")

        # 연도와 월 추출
        valid_data['completion_year'] = valid_data['Completion Date'].dt.year
        valid_data['completion_month'] = valid_data['Completion Date'].dt.month

        # 기준 연도 및 월 필터링
        date_filtered = valid_data[
            (valid_data['completion_year'] == ANALYSIS_YEAR) &
            (valid_data['completion_month'] <= ANALYSIS_MONTH)
        ].copy()

        logger.info(f"  - 기준월 필터링 결과:")
        logger.info(f"    입력: {len(valid_data)}개")
        logger.info(f"    출력: {len(date_filtered)}개 ({ANALYSIS_YEAR}년 1~{ANALYSIS_MONTH}월)")
        logger.info(f"    제외: {len(valid_data) - len(date_filtered)}개 (기준월 이후 또는 다른 연도)")

        if len(date_filtered) == 0:
            logger.warning(f"{ANALYSIS_YEAR}년 1~{ANALYSIS_MONTH}월 완료 데이터가 없습니다.")
            return None

        # 샘플 데이터 확인
        logger.info(f"  - 최종 필터링된 데이터 샘플:")
        logger.info(f"    Completion Date 샘플: {date_filtered['Completion Date'].head(3).tolist()}")
        logger.info(f"    Education Hours 샘플: {date_filtered['Education Hours'].head(3).tolist()}")
        logger.info(f"    Final Sub. 샘플: {date_filtered['Final Sub.'].head(5).tolist()}")

        # 법인명 공백 제거
        logger.info("4.2.3단계: 법인명 공백을 제거합니다...")
        logger.info("  - 방법: 앞뒤 공백만 제거 (대문자는 이미 통일되어 있음)")

        # Final Sub. 컬럼 공백 제거
        date_filtered['Final Sub. Normalized'] = date_filtered['Final Sub.'].str.strip()

        # 매핑 테이블 생성
        normalization_map = {}
        for original_name in date_filtered['Final Sub.'].unique():
            if pd.notna(original_name):
                normalized_name = str(original_name).strip()
                if normalized_name not in normalization_map:
                    normalization_map[normalized_name] = original_name

        logger.info(f"  - 공백 제거된 법인 수: {len(normalization_map)}개")

        # 누적 실제 수강 시간 계산
        logger.info(f"4.2.4단계: {ANALYSIS_YEAR}년 1~{ANALYSIS_MONTH}월 누적 실제 수강 시간을 계산합니다...")
        logger.info(f"  - 계산 방법:")
        logger.info(f"    법인별로 Education Hours 합산")
        logger.info(f"    합계가 0보다 큰 법인만 분석 대상")

        # 법인별 누적 Education Hours 합계
        subsidiary_actual_totals = date_filtered.groupby('Final Sub. Normalized')['Education Hours'].sum().sort_values(ascending=False)

        logger.info(f"  - 누적 실제 수강 시간 계산 완료")
        logger.info(f"  - 분석 법인: {len(subsidiary_actual_totals)}개")
        logger.info("    의미: {ANALYSIS_YEAR}년 1~{ANALYSIS_MONTH}월 동안 교육을 완료한 법인 수")
        logger.info(f"    전체 법인(99개) 중 누적 교육 완료 기록이 있는 법인: {len(subsidiary_actual_totals)}개")
        logger.info(f"    나머지 {99 - len(subsidiary_actual_totals)}개 법인: 누적 완료된 교육이 0 또는 데이터 없음")

        # 법인별 누적 시간 샘플 출력
        if len(subsidiary_actual_totals) > 0:
            logger.info("  - 법인별 누적 실제 수강 시간 샘플 (상위 10개):")
            for subsidiary, hours in subsidiary_actual_totals.head(10).items():
                original_name = normalization_map.get(subsidiary, subsidiary)
                logger.info(f"    {original_name}: {hours:.2f}시간")

        # 결과 요약
        logger.info("4.2.5단계: 결과를 요약합니다...")

        total_hours = subsidiary_actual_totals.sum()
        avg_hours = subsidiary_actual_totals.mean()
        max_hours = subsidiary_actual_totals.max()
        min_hours = subsidiary_actual_totals.min()

        logger.info(f"  - 전체 누적 실제 수강 시간: {total_hours:.2f}시간")
        logger.info(f"  - 법인당 평균 수강 시간: {avg_hours:.2f}시간")
        logger.info(f"  - 최대 수강 시간: {max_hours:.2f}시간")
        logger.info(f"  - 최소 수강 시간: {min_hours:.2f}시간")

        logger.info(f"✓ 4.2단계 완료: {ANALYSIS_YEAR}년 1~{ANALYSIS_MONTH}월 누적 실제 수강 시간 계산 완료")

        return {
            'subsidiary_actual_totals': subsidiary_actual_totals,
            'total_hours': total_hours,
            'avg_hours': avg_hours,
            'max_hours': max_hours,
            'min_hours': min_hours,
            'total_subsidiaries': len(subsidiary_actual_totals),
            'normalization_map': normalization_map
        }

    except Exception as e:
        logger.error(f"✗ 오류 발생: {e}")
        return None

def calculate_monthly_completion_rate(monthly_learning_result, monthly_actual_result):
    """
    법인별 누적 이수율을 계산하는 함수 (4.3단계)
    """
    try:
        # 4.3단계: 법인별 누적 이수율 계산
        logger.info("=" * 80)
        logger.info(f"4.3단계: {ANALYSIS_YEAR}년 1~{ANALYSIS_MONTH}월 법인별 누적 이수율을 계산합니다...")
        logger.info("=" * 80)
        logger.info("  - 계산 방법:")
        logger.info("    이수율 = (실제 수강 시간 / 계획 시간) × 100")
        logger.info("    조건:")
        logger.info("      1) 계획 시간 > 0, 실제 시간 > 0: 정상 계산")
        logger.info("      2) 계획만 있음: 0%")
        logger.info("      3) 완료만 있음: 0%")
        logger.info("-" * 80)

        if monthly_learning_result is None or monthly_actual_result is None:
            logger.error("✗ 4.1단계 또는 4.2단계 결과가 없습니다.")
            return None

        # 4.1단계와 4.2단계에서 법인별 누적 시간 가져오기
        plan_totals = monthly_learning_result['subsidiary_totals']  # Series
        actual_totals = monthly_actual_result['subsidiary_actual_totals']  # Series

        # 모든 법인 목록 수집
        all_subsidiaries = set(plan_totals.index) | set(actual_totals.index)

        logger.info(f"  - 계획 데이터 법인 수: {len(plan_totals)}개")
        logger.info(f"  - 실제 데이터 법인 수: {len(actual_totals)}개")
        logger.info(f"  - 전체 법인 수 (합집합): {len(all_subsidiaries)}개")



        # 법인별 누적 이수율 계산
        logger.info(f"4.3.1단계: 법인별 {ANALYSIS_YEAR}년 1~{ANALYSIS_MONTH}월 누적 이수율을 계산합니다...")

        completion_rates = {}
        matched_subsidiaries = []
        planned_only = []
        completed_only = []

        for subsidiary in all_subsidiaries:
            # 계획 시간 가져오기
            plan_hrs = plan_totals.get(subsidiary, 0)

            # 실제 시간 가져오기
            actual_hrs = actual_totals.get(subsidiary, 0)

            # 이수율 계산
            if plan_hrs > 0 and actual_hrs > 0:
                # 정상 계산
                rate = (actual_hrs / plan_hrs) * 100
                rate = math.ceil(rate * 100) / 100  # 소수점 2번째에서 올림
                matched_subsidiaries.append(subsidiary)
                logger.info(f"  {subsidiary}: 계획 {plan_hrs:.2f}시간, 실제 {actual_hrs:.2f}시간 → {rate:.2f}%")
            elif plan_hrs > 0 and actual_hrs == 0:
                # 계획만 있음 → 0%
                rate = 0.0
                planned_only.append(subsidiary)
                logger.info(f"  {subsidiary}: 계획 {plan_hrs:.2f}시간, 실제 0시간 → 0.00% (계획만 있음)")
            elif plan_hrs == 0 and actual_hrs > 0:
                # 완료만 있음 → 0%
                rate = 0.0
                completed_only.append(subsidiary)
                logger.info(f"  {subsidiary}: 계획 0시간, 실제 {actual_hrs:.2f}시간 → 0.00% (완료만 있음)")
            else:
                # 둘 다 0
                rate = 0.0

            completion_rates[subsidiary] = {
                'rate': rate,
                'plan_hrs': plan_hrs,
                'actual_hrs': actual_hrs
            }

        # 결과 요약
        logger.info("4.3.2단계: 결과를 요약합니다...")
        logger.info(f"  - 매칭 법인: {len(matched_subsidiaries)}개 (계획 & 실제 모두 있음)")
        logger.info(f"  - 계획만 있는 법인: {len(planned_only)}개")
        logger.info(f"  - 완료만 있는 법인: {len(completed_only)}개")

        if planned_only:
            logger.info(f"  - 계획만 있는 법인 목록: {planned_only[:5]}...")
        if completed_only:
            logger.info(f"  - 완료만 있는 법인 목록: {completed_only[:5]}...")

        # 이수율 통계
        valid_rates = [data['rate'] for data in completion_rates.values() if data['rate'] > 0]
        if valid_rates:
            avg_rate = sum(valid_rates) / len(valid_rates)
            max_rate = max(valid_rates)
            min_rate = min(valid_rates)
            logger.info(f"  - 평균 이수율: {avg_rate:.2f}%")
            logger.info(f"  - 최대 이수율: {max_rate:.2f}%")
            logger.info(f"  - 최소 이수율: {min_rate:.2f}%")

        logger.info(f"✓ 4.3단계 완료: {ANALYSIS_YEAR}년 1~{ANALYSIS_MONTH}월 누적 이수율 계산 완료")

        return {
            'completion_rates': completion_rates,
            'matched_subsidiaries': matched_subsidiaries,
            'planned_only': planned_only,
            'completed_only': completed_only,
            'total_subsidiaries': len(all_subsidiaries)
        }

    except Exception as e:
        logger.error(f"✗ 오류 발생: {e}")
        return None

def calculate_new_hire_completion_rate(join_table):
    """
    신입사원 교육 이수율을 계산하는 함수 (5단계)
    """
    try:
        # 5단계: 신입사원 교육 이수율 계산
        logger.info("=" * 80)
        logger.info("5단계: 신입사원 교육 이수율을 계산합니다...")
        logger.info("=" * 80)
        logger.info("사용 파일: join_hr_lms.csv")
        logger.info(f"전체 레코드 수: {len(join_table)}개")
        logger.info("-" * 80)

        # 필요한 컬럼 확인
        required_columns = ['New Hire', 'Course name', 'Item ID', 'Completion status', 'Hire Date', 'Staff/Operator', 'Position']
        missing_columns = [col for col in required_columns if col not in join_table.columns]

        if missing_columns:
            logger.error(f"✗ 필요한 컬럼이 없습니다: {missing_columns}")
            return None

        logger.info(f"✓ 필요한 컬럼 확인 완료: {required_columns}")

        # 5.1단계: 신입사원 필터링
        logger.info("5.1단계: 신입사원을 필터링합니다...")
        logger.info("  - 필터링 조건 (AND 조건):")
        logger.info("    1) New Hire = 'Y'")
        logger.info("    2) Staff/Operator = 'Staff'")
        logger.info("    3) Position = null (관리자 제외)")
        logger.info("  - 목적: 신입 Staff만 대상으로 교육 이수율 계산")

        # 신입사원 조건: New Hire가 'Y' AND Staff/Operator가 'Staff' AND Position이 null
        new_hire_condition = (
            (join_table['New Hire'] == 'Y') &
            (join_table['Staff/Operator'] == 'Staff') &
            (join_table['Position'].isna())
        )

        new_hire_employees = join_table[new_hire_condition].copy()

        # 각 조건별 인원 수 확인
        new_hire_y_count = len(join_table[join_table['New Hire'] == 'Y'])
        staff_count = len(join_table[join_table['Staff/Operator'] == 'Staff'])
        null_position_count = len(join_table[join_table['Position'].isna()])

        logger.info(f"  - 조건별 레코드 수:")
        logger.info(f"    New Hire = 'Y': {new_hire_y_count}개")
        logger.info(f"    Staff/Operator = 'Staff': {staff_count}개")
        logger.info(f"    Position = null: {null_position_count}개")
        logger.info(f"  - 최종 필터링 결과: {len(new_hire_employees)}개 (3개 조건 모두 만족)")

        if len(new_hire_employees) == 0:
            logger.warning("신입사원이 없습니다.")
            return None

        # 5.2단계: 신입사원 교육 과정 필터링
        logger.info("5.2단계: 신입사원 교육 과정을 필터링합니다...")
        logger.info("  - 필터링 조건 (OR 조건):")
        logger.info("    Course name에 다음 키워드 포함 (대소문자 구분 없음):")

        # 신입사원 교육 과정 조건
        course_keywords = ['new comer', 'new employee', 'new joiner', 'new member', 'new lger', 'orientation', '新入']

        for idx, keyword in enumerate(course_keywords, 1):
            logger.info(f"      {idx}) '{keyword}'")
        logger.info("    또는 Item ID = '20015'")
        logger.info("  - 목적: 신입사원 교육 과정만 선별")

        # 각 키워드별 개수 확인
        logger.info("  - 키워드별 레코드 수:")
        for keyword in course_keywords:
            keyword_count = len(new_hire_employees[new_hire_employees['Course name'].str.contains(keyword, case=False, na=False)])
            if keyword_count > 0:
                logger.info(f"    '{keyword}': {keyword_count}개")

        # Item ID 조건 확인
        item_id_count = len(new_hire_employees[new_hire_employees['Item ID'] == "20015"])
        logger.info(f"    Item ID = '20015': {item_id_count}개")

        # Course name에 키워드가 포함되거나 Item ID가 20015인 경우
        course_condition = (
            new_hire_employees['Course name'].str.contains('|'.join(course_keywords), case=False, na=False) |
            (new_hire_employees['Item ID'] == "20015")
        )

        new_hire_courses = new_hire_employees[course_condition].copy()

        logger.info(f"  - 최종 필터링 결과: {len(new_hire_courses)}개 레코드 (신입사원 교육 과정)")

        # 5.3단계: 직원별 이수 상태 계산
        logger.info("5.3단계: 직원별 이수 상태를 계산합니다...")
        logger.info(f"  - 이수 상태 판단 기준 ({ANALYSIS_YEAR}년 {ANALYSIS_MONTH}월 1일 기준):")
        logger.info("    1) '이수' (C): Completion status가 '-C'로 끝나는 과정이 1개 이상")
        logger.info("    2) '보류' (H): 이수 과정 없음 AND 입사일부터 3개월 미만")
        logger.info("    3) '미이수' (N): 이수 과정 없음 AND 입사일부터 3개월 이상")
        logger.info("  - 목적: 신입사원별 교육 이수 여부 확인")

        # Hire Date를 datetime으로 변환
        new_hire_courses['Hire Date'] = pd.to_datetime(new_hire_courses['Hire Date'], errors='coerce')

        # 분석 기준 날짜 (8월 1일)
        analysis_date = pd.Timestamp(f"{ANALYSIS_YEAR}-{ANALYSIS_MONTH:02d}-01")
        logger.info(f"  - 분석 기준 날짜: {analysis_date.strftime('%Y-%m-%d')}")

        # 직원별로 그룹화하여 이수 상태 계산 (Employee Number 기준)
        employee_completion = {}

        # Employee Number로 그룹핑 (고유 직원별로 처리)
        unique_employees = new_hire_courses['Employee Number'].unique()
        logger.info(f"  - 고유 신입사원 수: {len(unique_employees)}명")

        for emp_no in unique_employees:
            emp_data = new_hire_courses[new_hire_courses['Employee Number'] == emp_no]

            # 해당 직원의 신입사원 교육 수료 여부 확인
            completed_courses = emp_data[emp_data['Completion status'].str.endswith('-C', na=False)]

            # Hire Date 가져오기 (첫 번째 레코드에서)
            hire_date = emp_data['Hire Date'].iloc[0]

            # 입사 후 경과 개월 수 계산
            if pd.notna(hire_date):
                months_diff = (analysis_date.year - hire_date.year) * 12 + (analysis_date.month - hire_date.month)
            else:
                months_diff = 999  # Hire Date가 없으면 충분히 큰 값으로 설정

            # 이수 상태 결정
            if len(completed_courses) > 0:
                status = 'C'  # 이수
            elif months_diff < 3:
                status = 'H'  # 보류 (입사 3개월 미만)
            else:
                status = 'N'  # 미이수 (입사 3개월 이상)

            employee_completion[emp_no] = {
                'status': status,
                'hire_date': hire_date,
                'months_since_hire': months_diff,
                'completed_courses': len(completed_courses)
            }

        # 5.4단계: 전체 결과 요약
        logger.info("5.4단계: 전체 결과를 요약합니다...")
        logger.info("  - 이수율 계산 공식:")
        logger.info("    이수율 = (이수 + 보류) / 전체 × 100")
        logger.info("    이유: 입사 3개월 미만 신입은 '보류'로 분류하여 이수로 간주")

        status_counts = {'C': 0, 'N': 0, 'H': 0}
        for emp_data in employee_completion.values():
            status_counts[emp_data['status']] += 1

        logger.info("  - 신입사원 교육 이수 현황 (전체):")
        logger.info(f"    이수(C): {status_counts['C']}명")
        logger.info(f"    미이수(N): {status_counts['N']}명")
        logger.info(f"    보류(H): {status_counts['H']}명 (입사 3개월 미만)")

        total_new_hire = len(employee_completion)
        # 이수율 = (이수 + 보류) / 전체 * 100
        eligible_count = status_counts['C'] + status_counts['H']
        completion_rate = (eligible_count / total_new_hire * 100) if total_new_hire > 0 else 0
        logger.info(f"  - 전체 이수율: {completion_rate:.2f}%")
        logger.info(f"    계산: ({status_counts['C']} + {status_counts['H']}) / {total_new_hire} × 100")

        # 5.5단계: 법인별 결과 계산
        logger.info("5.5단계: 법인별 결과를 계산합니다...")
        logger.info("  - 계산 방법:")
        logger.info("    법인별 이수율 = (법인 이수 + 법인 보류) / 법인 전체 × 100")
        logger.info("  - 특수 케이스 처리:")
        logger.info("    신입사원이 없는 법인: 0%로 처리")

        # 직원별 법인 정보 추가 (Employee Number로 그룹핑)
        subsidiary_completion = {}

        for emp_no, emp_data in employee_completion.items():
            # 해당 직원의 법인 정보 가져오기
            emp_row = new_hire_courses[new_hire_courses['Employee Number'] == emp_no].iloc[0]
            subsidiary = emp_row['Final Sub.']

            if subsidiary not in subsidiary_completion:
                subsidiary_completion[subsidiary] = {'C': 0, 'N': 0, 'H': 0, 'total': 0}

            subsidiary_completion[subsidiary][emp_data['status']] += 1
            subsidiary_completion[subsidiary]['total'] += 1

        # 법인별 결과 출력 (이수율 높은 순으로 정렬)
        logger.info("  - 법인별 신입사원 교육 이수 현황:")

        subsidiary_rates = []
        subsidiaries_with_new_hires = []
        subsidiaries_without_new_hires = []

        for subsidiary, counts in subsidiary_completion.items():
            # 이수율 = (이수 + 보류) / 전체 * 100
            eligible_count = counts['C'] + counts['H']
            if counts['total'] > 0:
                rate = (eligible_count / counts['total'] * 100)
                subsidiary_rates.append((subsidiary, rate, counts, eligible_count))
                subsidiaries_with_new_hires.append(subsidiary)
            else:
                # 신입사원이 없는 경우 → 0%
                rate = 0.0
                subsidiary_rates.append((subsidiary, rate, counts, eligible_count))
                subsidiaries_without_new_hires.append(subsidiary)

        # 신입사원 없는 법인 정보
        if subsidiaries_without_new_hires:
            logger.info(f"  - 신입사원이 없는 법인: {len(subsidiaries_without_new_hires)}개")
            logger.info(f"    법인 목록: {', '.join(subsidiaries_without_new_hires)}")
            logger.info("    처리: 모두 0%로 계산")
        else:
            logger.info("  - 신입사원이 없는 법인: 없음")

        # 이수율 높은 순으로 정렬
        subsidiary_rates.sort(key=lambda x: x[1], reverse=True)

        logger.info(f"  - 신입사원이 있는 법인: {len(subsidiaries_with_new_hires)}개")
        for subsidiary, rate, counts, eligible_count in subsidiary_rates:
            if counts['total'] > 0:
                logger.info(f"    {subsidiary}: {rate:.2f}% (이수: {counts['C']}명, 미이수: {counts['N']}명, 보류: {counts['H']}명)")

        logger.info(f"  - 총 {len(subsidiary_rates)}개 법인 분석 완료")

        # 샘플 결과 출력 (상위 5개)
        logger.info("  - 신입사원 이수 상태 샘플 (5명):")
        sample_count = 0
        for emp_no, emp_data in employee_completion.items():
            if sample_count >= 5:
                break
            hire_date_str = emp_data['hire_date'].strftime('%Y-%m-%d') if pd.notna(emp_data['hire_date']) else 'N/A'
            status_label = {'C': '이수', 'N': '미이수', 'H': '보류'}.get(emp_data['status'], emp_data['status'])
            logger.info(f"    사번 {emp_no}: {status_label} (입사: {hire_date_str}, 경과: {emp_data['months_since_hire']}개월, 수료: {emp_data['completed_courses']}개)")
            sample_count += 1

        logger.info(f"✓ 5단계 완료: 신입사원 교육 이수율 계산 완료")

        return {
            'employee_completion': employee_completion,
            'status_counts': status_counts,
            'total_new_hire': total_new_hire,
            'completion_rate': completion_rate,
            'subsidiary_completion': subsidiary_completion
        }

    except Exception as e:
        logger.error(f"✗ 오류 발생: {e}")
        return None

def calculate_hipo_completion_rate(join_table):
    """
    6단계: 핵심인재 교육 이수율 계산
    """
    try:
        logger.info("=" * 80)
        logger.info("6단계: 핵심인재 교육 이수율을 계산합니다...")
        logger.info("=" * 80)
        logger.info("사용 파일: join_hr_lms.csv")
        logger.info(f"전체 레코드 수: {len(join_table)}개")
        logger.info("-" * 80)

        # 필요한 컬럼 확인
        required_columns = ['HIPO Type', 'Completion status', 'Employee Number', 'Staff/Operator']
        missing_columns = [col for col in required_columns if col not in join_table.columns]

        if missing_columns:
            logger.error(f"✗ 필요한 컬럼이 없습니다: {missing_columns}")
            return None

        logger.info(f"✓ 필요한 컬럼 확인 완료: {required_columns}")

        # 6.1단계: EIP 직원 필터링
        logger.info("6.1단계: EIP 핵심인재를 필터링합니다...")
        logger.info("  - 필터링 조건 (AND 조건):")
        logger.info("    1) HIPO Type = 'EIP'")
        logger.info("    2) Staff/Operator = 'Staff'")
        logger.info("  - 목적: EIP 핵심인재만 대상으로 교육 이수율 계산")

        # HIPO Type이 EIP이고 Staff인 직원들 필터링
        hipo_type_count = len(join_table[join_table['HIPO Type'].notna()])
        eip_count = len(join_table[join_table['HIPO Type'] == 'EIP'])
        staff_count = len(join_table[join_table['Staff/Operator'] == 'Staff'])

        eip_condition = (join_table['HIPO Type'] == 'EIP') & (join_table['Staff/Operator'] == 'Staff')
        eip_employees = join_table[eip_condition].copy()

        logger.info(f"  - 조건별 레코드 수:")
        logger.info(f"    HIPO Type not null: {hipo_type_count}개")
        logger.info(f"    HIPO Type = 'EIP': {eip_count}개")
        logger.info(f"    Staff/Operator = 'Staff': {staff_count}개")
        logger.info(f"  - 최종 필터링 결과: {len(eip_employees)}개 (2개 조건 모두 만족)")

        # 고유한 EIP 직원 수 계산
        unique_eip_employees = eip_employees['Employee Number'].nunique()
        logger.info(f"  - 고유 EIP 핵심인재 수: {unique_eip_employees}명")

        if len(eip_employees) > unique_eip_employees:
            avg_records_per_employee = len(eip_employees) / unique_eip_employees
            logger.info(f"  - 평균 직원당 교육 레코드 수: {avg_records_per_employee:.1f}개")

        # 6.1.1단계: EIP 직원별 이수 상태 확인
        logger.info("6.1.1단계: 각 EIP 직원의 교육 이수 상태를 계산합니다...")
        logger.info("  - 이수 상태 판단 기준:")
        logger.info("    1) '이수' (C): Completion status가 '-C'로 끝나는 과정이 1개 이상")
        logger.info("    2) '미이수' (N): 이수 과정 없음")
        logger.info("  - 목적: EIP 핵심인재별 교육 이수 여부 확인")

        # Employee Number로 그룹핑하여 각 직원의 이수 상태 확인
        eip_completion = {}
        completion_details = {}

        for emp_no in eip_employees['Employee Number'].unique():
            emp_data = eip_employees[eip_employees['Employee Number'] == emp_no]

            # 해당 직원의 Completion status 중에 '-C'로 끝나는 것이 있는지 확인
            completed_courses = emp_data[emp_data['Completion status'].str.endswith('-C', na=False)]
            total_courses = len(emp_data)

            if len(completed_courses) > 0:
                eip_completion[emp_no] = '이수'
                completion_details[emp_no] = {
                    'total_courses': total_courses,
                    'completed_courses': len(completed_courses),
                    'completion_rate': len(completed_courses) / total_courses * 100
                }
            else:
                eip_completion[emp_no] = '미이수'
                completion_details[emp_no] = {
                    'total_courses': total_courses,
                    'completed_courses': 0,
                    'completion_rate': 0.0
                }

        # EIP 이수 상태별 통계
        logger.info("6.1.2단계: EIP 전체 결과를 요약합니다...")
        logger.info("  - 이수율 계산 공식:")
        logger.info("    이수율 = 이수 / 전체 × 100")

        completed_employees = [emp for emp, status in eip_completion.items() if status == '이수']
        incomplete_employees = [emp for emp, status in eip_completion.items() if status == '미이수']

        logger.info(f"  - EIP 핵심인재 교육 이수 현황:")
        logger.info(f"    이수: {len(completed_employees)}명")
        logger.info(f"    미이수: {len(incomplete_employees)}명")
        logger.info(f"    전체: {unique_eip_employees}명")

        # EIP 이수율 계산
        eip_completion_rate = (len(completed_employees) / unique_eip_employees * 100) if unique_eip_employees > 0 else 0
        logger.info(f"  - EIP 전체 이수율: {eip_completion_rate:.2f}%")
        logger.info(f"    계산: {len(completed_employees)} / {unique_eip_employees} × 100")

        # 6.2단계: GLP 핵심인재 필터링
        logger.info("6.2단계: GLP 핵심인재를 필터링합니다...")
        logger.info("  - 필터링 조건 (AND 조건):")
        logger.info("    1) HIPO Type = 'GLP'")
        logger.info("    2) Staff/Operator = 'Staff'")
        logger.info("  - 목적: GLP 핵심인재만 대상으로 교육 이수율 계산")

        # HIPO Type이 GLP이고 Staff인 직원들 필터링
        glp_count = len(join_table[join_table['HIPO Type'] == 'GLP'])

        glp_condition = (join_table['HIPO Type'] == 'GLP') & (join_table['Staff/Operator'] == 'Staff')
        glp_employees = join_table[glp_condition].copy()

        logger.info(f"  - 조건별 레코드 수:")
        logger.info(f"    HIPO Type = 'GLP': {glp_count}개")
        logger.info(f"    Staff/Operator = 'Staff': {staff_count}개")
        logger.info(f"  - 최종 필터링 결과: {len(glp_employees)}개 (2개 조건 모두 만족)")

        # 고유한 GLP 직원 수 계산
        unique_glp_employees = glp_employees['Employee Number'].nunique()
        logger.info(f"  - 고유 GLP 핵심인재 수: {unique_glp_employees}명")

        # 중복 분석
        if len(glp_employees) > unique_glp_employees:
            avg_records_per_employee = len(glp_employees) / unique_glp_employees
            logger.info(f"  - 평균 직원당 교육 레코드 수: {avg_records_per_employee:.1f}개")

        # Employee Number로 그룹핑하여 각 직원의 이수 상태 확인
        glp_completion = {}
        glp_completion_details = {}

        logger.info("6.2.1단계: 각 GLP 직원의 교육 이수 상태를 계산합니다...")
        logger.info("  - 이수 상태 판단 기준:")
        logger.info("    1) '이수' (C): Completion status가 '-C'로 끝나는 과정이 1개 이상")
        logger.info("    2) '미이수' (N): 이수 과정 없음")
        logger.info("  - 목적: GLP 핵심인재별 교육 이수 여부 확인")

        for emp_no in glp_employees['Employee Number'].unique():
            emp_data = glp_employees[glp_employees['Employee Number'] == emp_no]

            # 해당 직원의 Completion status 중에 '-C'로 끝나는 것이 있는지 확인
            completed_courses = emp_data[emp_data['Completion status'].str.endswith('-C', na=False)]
            total_courses = len(emp_data)

            if len(completed_courses) > 0:
                glp_completion[emp_no] = '이수'
                glp_completion_details[emp_no] = {
                    'total_courses': total_courses,
                    'completed_courses': len(completed_courses),
                    'completion_rate': len(completed_courses) / total_courses * 100
                }
            else:
                glp_completion[emp_no] = '미이수'
                glp_completion_details[emp_no] = {
                    'total_courses': total_courses,
                    'completed_courses': 0,
                    'completion_rate': 0.0
                }

        # 이수 상태별 통계
        logger.info("6.2.2단계: GLP 전체 결과를 요약합니다...")
        logger.info("  - 이수율 계산 공식:")
        logger.info("    이수율 = 이수 / 전체 × 100")

        glp_completed_employees = [emp for emp, status in glp_completion.items() if status == '이수']
        glp_incomplete_employees = [emp for emp, status in glp_completion.items() if status == '미이수']

        logger.info(f"  - GLP 핵심인재 교육 이수 현황:")
        logger.info(f"    이수: {len(glp_completed_employees)}명")
        logger.info(f"    미이수: {len(glp_incomplete_employees)}명")
        logger.info(f"    전체: {unique_glp_employees}명")

        # GLP 이수율 계산
        glp_completion_rate = (len(glp_completed_employees) / unique_glp_employees * 100) if unique_glp_employees > 0 else 0
        logger.info(f"  - GLP 전체 이수율: {glp_completion_rate:.2f}%")
        logger.info(f"    계산: {len(glp_completed_employees)} / {unique_glp_employees} × 100")

        # 6.3단계: 이수 상태 샘플
        logger.info("6.3단계: EIP/GLP 핵심인재 이수 상태 샘플...")

        # EIP 샘플
        logger.info("  - EIP 핵심인재 이수 상태 샘플 (5명):")
        sample_count = 0
        for emp_no, status in eip_completion.items():
            if sample_count >= 5:
                break
            logger.info(f"    사번 {emp_no}: {status}")
            sample_count += 1

        # GLP 샘플
        logger.info("  - GLP 핵심인재 이수 상태 샘플 (5명):")
        sample_count = 0
        for emp_no, status in glp_completion.items():
            if sample_count >= 5:
                break
            logger.info(f"    사번 {emp_no}: {status}")
            sample_count += 1

        # 6.4단계: 법인별 EIP, GLP 이수율 분석
        logger.info("6.4단계: 법인별 EIP, GLP 이수율을 분석합니다...")

        # EIP 법인별 분석
        logger.info("6.4.1단계: 법인별 EIP 이수율을 계산합니다...")

        eip_subsidiary_completion = {}

        for emp_no, status in eip_completion.items():
            # 해당 직원의 법인 정보 가져오기
            emp_row = eip_employees[eip_employees['Employee Number'] == emp_no].iloc[0]
            subsidiary = emp_row['Final Sub.']

            if subsidiary not in eip_subsidiary_completion:
                eip_subsidiary_completion[subsidiary] = {'이수': 0, '미이수': 0, '전체': 0}

            eip_subsidiary_completion[subsidiary][status] += 1
            eip_subsidiary_completion[subsidiary]['전체'] += 1

        # EIP 법인별 결과 출력 (이수율 높은 순으로 정렬)
        logger.info("법인별 EIP 핵심인재 교육 이수 현황:")

        eip_subsidiary_rates = []
        for subsidiary, counts in eip_subsidiary_completion.items():
            if counts['전체'] > 0:
                rate = (counts['이수'] / counts['전체'] * 100)
                eip_subsidiary_rates.append((subsidiary, rate, counts))

        # 이수율 높은 순으로 정렬
        eip_subsidiary_rates.sort(key=lambda x: x[1], reverse=True)

        for subsidiary, rate, counts in eip_subsidiary_rates:
            logger.info(f"  {subsidiary}: {rate:.2f}% (이수: {counts['이수']}명, 미이수: {counts['미이수']}명, 전체: {counts['전체']}명)")

        # GLP 법인별 분석
        logger.info("6.4.2단계: 법인별 GLP 이수율을 계산합니다...")

        glp_subsidiary_completion = {}

        for emp_no, status in glp_completion.items():
            # 해당 직원의 법인 정보 가져오기
            emp_row = glp_employees[glp_employees['Employee Number'] == emp_no].iloc[0]
            subsidiary = emp_row['Final Sub.']

            if subsidiary not in glp_subsidiary_completion:
                glp_subsidiary_completion[subsidiary] = {'이수': 0, '미이수': 0, '전체': 0}

            glp_subsidiary_completion[subsidiary][status] += 1
            glp_subsidiary_completion[subsidiary]['전체'] += 1

        # GLP 법인별 결과 출력 (이수율 높은 순으로 정렬)
        logger.info("법인별 GLP 핵심인재 교육 이수 현황:")

        glp_subsidiary_rates = []
        for subsidiary, counts in glp_subsidiary_completion.items():
            if counts['전체'] > 0:
                rate = (counts['이수'] / counts['전체'] * 100)
                glp_subsidiary_rates.append((subsidiary, rate, counts))

        # 이수율 높은 순으로 정렬
        glp_subsidiary_rates.sort(key=lambda x: x[1], reverse=True)

        for subsidiary, rate, counts in glp_subsidiary_rates:
            logger.info(f"  {subsidiary}: {rate:.2f}% (이수: {counts['이수']}명, 미이수: {counts['미이수']}명, 전체: {counts['전체']}명)")

        # EIP vs GLP 법인별 비교 요약
        logger.info("6.4.3단계: EIP vs GLP 법인별 비교 요약:")

        # 모든 법인 수집
        all_eip_subsidiaries = set(eip_subsidiary_completion.keys())
        all_glp_subsidiaries = set(glp_subsidiary_completion.keys())
        all_subsidiaries = all_eip_subsidiaries | all_glp_subsidiaries

        logger.info(f"✓ EIP 보유 법인: {len(all_eip_subsidiaries)}개")
        logger.info(f"✓ GLP 보유 법인: {len(all_glp_subsidiaries)}개")
        logger.info(f"✓ 전체 핵심인재 보유 법인: {len(all_subsidiaries)}개")

        # EIP와 GLP 모두 보유한 법인
        both_subsidiaries = all_eip_subsidiaries & all_glp_subsidiaries
        logger.info(f"✓ EIP와 GLP 모두 보유한 법인: {len(both_subsidiaries)}개")

        logger.info(f"✓ 6단계 완료: 핵심인재(EIP/GLP) 교육 이수율 계산 완료")

        # 이수/미이수 카운트 생성
        eip_completion_counts = {'이수': len(completed_employees), '미이수': len(incomplete_employees)}
        glp_completion_counts = {'이수': len(glp_completed_employees), '미이수': len(glp_incomplete_employees)}

        return {
            'eip_completion': eip_completion,
            'eip_completion_counts': eip_completion_counts,
            'total_eip': unique_eip_employees,
            'eip_completion_rate': eip_completion_rate,
            'eip_employees': eip_employees,
            'glp_completion': glp_completion,
            'glp_completion_counts': glp_completion_counts,
            'total_glp': unique_glp_employees,
            'glp_completion_rate': glp_completion_rate,
            'glp_employees': glp_employees
        }

    except Exception as e:
        logger.error(f"✗ 오류 발생: {e}")
        return None


def calculate_new_leader_completion_rate(join_table):
    """
    7단계: 신입 팀장 교육 이수율 계산
    """
    try:
        logger.info("=" * 80)
        logger.info("7단계: 신입 팀장 교육 이수율을 계산합니다...")
        logger.info("=" * 80)
        logger.info("사용 파일: join_hr_lms.csv")
        logger.info(f"전체 레코드 수: {len(join_table)}개")
        logger.info("-" * 80)

        # 필요한 컬럼 확인
        required_columns = ['New Leader', 'Item ID', 'Completion status', 'Employee Number', 'Staff/Operator']
        missing_columns = [col for col in required_columns if col not in join_table.columns]

        if missing_columns:
            logger.error(f"✗ 필요한 컬럼이 없습니다: {missing_columns}")
            logger.warning("  - New Leader 컬럼이 없을 수 있습니다. 모든 법인을 0%로 처리합니다.")
            return {
                'new_leader_completion': {},
                'new_leader_completion_counts': {'이수': 0, '미이수': 0},
                'total_new_leader': 0,
                'new_leader_completion_rate': 0,
                'new_leader_employees': pd.DataFrame()
            }

        logger.info(f"✓ 필요한 컬럼 확인 완료: {required_columns}")

        # 7.1단계: 신입 팀장 필터링
        logger.info("7.1단계: 신입 팀장을 필터링합니다...")
        logger.info("  - 필터링 조건 (AND 조건):")
        logger.info("    1) New Leader = 'Y'")
        logger.info("    2) Staff/Operator = 'Staff'")
        logger.info("  - 목적: 신입 팀장만 대상으로 교육 이수율 계산")

        # New Leader = 'Y' AND Staff/Operator = 'Staff'인 직원들 필터링
        new_leader_y_count = len(join_table[join_table['New Leader'] == 'Y'])
        staff_count = len(join_table[join_table['Staff/Operator'] == 'Staff'])

        new_leader_condition = (join_table['New Leader'] == 'Y') & (join_table['Staff/Operator'] == 'Staff')
        new_leader_employees = join_table[new_leader_condition].copy()

        logger.info(f"  - 조건별 레코드 수:")
        logger.info(f"    New Leader = 'Y': {new_leader_y_count}개")
        logger.info(f"    Staff/Operator = 'Staff': {staff_count}개")
        logger.info(f"  - 최종 필터링 결과: {len(new_leader_employees)}개 (2개 조건 모두 만족)")

        # 고유한 신입 팀장 수 계산
        unique_new_leader_employees = new_leader_employees['Employee Number'].nunique()
        logger.info(f"  - 고유 신입 팀장 수: {unique_new_leader_employees}명")

        if len(new_leader_employees) > unique_new_leader_employees:
            avg_records_per_employee = len(new_leader_employees) / unique_new_leader_employees
            logger.info(f"  - 평균 직원당 교육 레코드 수: {avg_records_per_employee:.1f}개")

        # 신입 팀장이 없는 경우
        if unique_new_leader_employees == 0:
            logger.warning("✗ 신입 팀장이 없습니다.")
            logger.info("  - 모든 법인의 New_Leader_Completion_Rate를 0%로 처리합니다.")
            return {
                'new_leader_completion': {},
                'new_leader_completion_counts': {'이수': 0, '미이수': 0},
                'total_new_leader': 0,
                'new_leader_completion_rate': 0,
                'new_leader_employees': new_leader_employees
            }

        # 7.2단계: 신입 팀장 교육 과정 필터링
        logger.info("7.2단계: 신입 팀장 교육 과정을 필터링합니다...")
        logger.info("  - 필터링 조건:")
        logger.info("    Item ID = '[LGE_HQ_Assimilation Workshop]'")
        logger.info("  - 목적: 신입 팀장 교육 과정만 선별")

        # Item ID = '[LGE_HQ_Assimilation Workshop]'인 레코드 필터링
        item_id_condition = new_leader_employees['Item ID'] == '[LGE_HQ_Assimilation Workshop]'
        new_leader_courses = new_leader_employees[item_id_condition].copy()

        logger.info(f"  - Item ID = '[LGE_HQ_Assimilation Workshop]': {len(new_leader_courses)}개 레코드")
        logger.info(f"  - 해당 과정 수강 중인 신입 팀장: {new_leader_courses['Employee Number'].nunique()}명")

        # 7.3단계: 직원별 이수 상태 계산
        logger.info("7.3단계: 직원별 이수 상태를 계산합니다...")
        logger.info("  - 이수 상태 판단 기준:")
        logger.info("    1) '이수' (C): Item ID = '[LGE_HQ_Assimilation Workshop]' 과정에서 Completion status가 '-C'로 끝남")
        logger.info("    2) '미이수' (N): 이수 과정 없음 (= 전체 - 이수)")
        logger.info("  - 목적: 신입 팀장별 교육 이수 여부 확인")

        # Employee Number로 그룹핑하여 각 직원의 이수 상태 확인
        new_leader_completion = {}
        completion_details = {}

        # 전체 신입 팀장 리스트 (과정 수강 여부와 무관)
        all_new_leaders = new_leader_employees['Employee Number'].unique()

        for emp_no in all_new_leaders:
            # 해당 직원의 신입 팀장 교육 과정 데이터
            emp_courses = new_leader_courses[new_leader_courses['Employee Number'] == emp_no]

            if len(emp_courses) == 0:
                # 과정 수강 기록이 없으면 미이수
                new_leader_completion[emp_no] = '미이수'
                completion_details[emp_no] = {
                    'total_courses': 0,
                    'completed_courses': 0,
                    'completion_rate': 0.0
                }
            else:
                # Completion status가 '-C'로 끝나는 것이 있는지 확인
                completed_courses = emp_courses[emp_courses['Completion status'].str.endswith('-C', na=False)]
                total_courses = len(emp_courses)

                if len(completed_courses) > 0:
                    new_leader_completion[emp_no] = '이수'
                    completion_details[emp_no] = {
                        'total_courses': total_courses,
                        'completed_courses': len(completed_courses),
                        'completion_rate': len(completed_courses) / total_courses * 100
                    }
                else:
                    new_leader_completion[emp_no] = '미이수'
                    completion_details[emp_no] = {
                        'total_courses': total_courses,
                        'completed_courses': 0,
                        'completion_rate': 0.0
                    }

        # 7.4단계: 전체 결과 요약
        logger.info("7.4단계: 신입 팀장 전체 결과를 요약합니다...")
        logger.info("  - 이수율 계산 공식:")
        logger.info("    이수율 = 이수 / 전체 × 100")

        completed_employees = [emp for emp, status in new_leader_completion.items() if status == '이수']
        incomplete_employees = [emp for emp, status in new_leader_completion.items() if status == '미이수']

        logger.info(f"  - 신입 팀장 교육 이수 현황:")
        logger.info(f"    이수: {len(completed_employees)}명")
        logger.info(f"    미이수: {len(incomplete_employees)}명")
        logger.info(f"    전체: {unique_new_leader_employees}명")

        # 신입 팀장 이수율 계산
        new_leader_completion_rate = (len(completed_employees) / unique_new_leader_employees * 100) if unique_new_leader_employees > 0 else 0
        logger.info(f"  - 신입 팀장 전체 이수율: {new_leader_completion_rate:.2f}%")
        logger.info(f"    계산: {len(completed_employees)} / {unique_new_leader_employees} × 100")

        # 7.5단계: 이수 상태 샘플
        logger.info("7.5단계: 신입 팀장 이수 상태 샘플 (최대 10명)...")
        sample_count = 0
        for emp_no, status in new_leader_completion.items():
            if sample_count >= 10:
                break
            logger.info(f"  사번 {emp_no}: {status}")
            sample_count += 1

        # 7.6단계: 법인별 신입 팀장 이수율 분석
        logger.info("7.6단계: 법인별 신입 팀장 이수율을 분석합니다...")

        new_leader_subsidiary_completion = {}

        for emp_no, status in new_leader_completion.items():
            # 해당 직원의 법인 정보 가져오기
            emp_row = new_leader_employees[new_leader_employees['Employee Number'] == emp_no].iloc[0]
            subsidiary = emp_row['Final Sub.']

            if subsidiary not in new_leader_subsidiary_completion:
                new_leader_subsidiary_completion[subsidiary] = {'이수': 0, '미이수': 0, '전체': 0}

            new_leader_subsidiary_completion[subsidiary][status] += 1
            new_leader_subsidiary_completion[subsidiary]['전체'] += 1

        # 법인별 결과 출력 (이수율 높은 순으로 정렬)
        logger.info("법인별 신입 팀장 교육 이수 현황:")

        new_leader_subsidiary_rates = []
        for subsidiary, counts in new_leader_subsidiary_completion.items():
            if counts['전체'] > 0:
                rate = (counts['이수'] / counts['전체'] * 100)
                new_leader_subsidiary_rates.append((subsidiary, rate, counts))

        # 이수율 높은 순으로 정렬
        new_leader_subsidiary_rates.sort(key=lambda x: x[1], reverse=True)

        for subsidiary, rate, counts in new_leader_subsidiary_rates:
            logger.info(f"  {subsidiary}: {rate:.2f}% (이수: {counts['이수']}명, 미이수: {counts['미이수']}명, 전체: {counts['전체']}명)")

        logger.info(f"✓ 7단계 완료: 신입 팀장 교육 이수율 계산 완료")

        # 이수/미이수 카운트 생성
        new_leader_completion_counts = {'이수': len(completed_employees), '미이수': len(incomplete_employees)}

        return {
            'new_leader_completion': new_leader_completion,
            'new_leader_completion_counts': new_leader_completion_counts,
            'total_new_leader': unique_new_leader_employees,
            'new_leader_completion_rate': new_leader_completion_rate,
            'new_leader_employees': new_leader_employees
        }

    except Exception as e:
        logger.error(f"✗ 오류 발생: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return None


def create_final_logic_data(step3_result, step4_result, step5_result, step6_result, step7_result, file_directory):
    """
    8단계: 최종 데이터 생성

    Args:
        step3_result: 3단계 결과
        step4_result: 4단계 결과
        step5_result: 5단계 결과
        step6_result: 6단계 결과
        step7_result: 7단계 결과
        file_directory (str): 파일 저장 디렉토리
    """
    try:
        logger.info("=" * 80)
        logger.info("8단계: 최종 데이터를 생성합니다...")
        logger.info("=" * 80)
        logger.info("입력 데이터:")
        logger.info("  - 3단계 결과: 법인별 과정 완료율")
        logger.info("  - 4단계 결과: 법인별 누적 이수율")
        logger.info("  - 5단계 결과: 신입사원 교육 이수율")
        logger.info("  - 6단계 결과: 핵심인재(EIP/GLP) 교육 이수율")
        logger.info("  - 7단계 결과: 신입 팀장 교육 이수율")
        logger.info(f"출력 파일: {file_directory}/logic.csv")
        logger.info("-" * 80)

        # 8.1단계: 3단계 결과를 logic.csv로 저장
        logger.info("8.1단계: 3단계 결과를 logic.csv로 저장합니다...")

        if step3_result is not None:
            # 3단계 결과에서 법인별 데이터 추출
            completion_data = step3_result.get('final_list', [])

            if completion_data:
                # DataFrame 생성
                logic_df = pd.DataFrame(completion_data)

                # 컬럼명 정리 및 데이터 변환
                logic_df = logic_df[['subsidiary', 'planned', 'completed', 'completion_rate']].copy()
                logic_df.columns = ['Subsidiary', 'Planned_Courses', 'Completed_Courses', 'Course_Completion_Rate']

                # 완료율은 이미 숫자로 되어 있음

                # 7.1.1단계: 4단계 결과와 조인
                logger.info("8.1.1단계: 4단계 결과(법인별 누적 이수율)와 조인합니다...")
                logger.info("  - 조인 방법:")
                logger.info("    조인 타입: OUTER JOIN")
                logger.info("    조인 키: Subsidiary (대소문자 구분 없음)")
                logger.info("    Left: 3단계 결과 (법인별 과정 완료율)")
                logger.info("    Right: 4단계 결과 (법인별 누적 이수율)")
                logger.info("  - OUTER JOIN 이유:")
                logger.info("    3단계에만 있는 법인: 과정 완료는 있지만 시간 이수율 데이터 없음")
                logger.info("    4단계에만 있는 법인: 시간 이수율은 있지만 과정 완료 데이터 없음")
                logger.info("    양쪽 모두 있는 법인: 모든 데이터 결합")

                if step4_result is not None:
                    # 4단계 결과에서 법인별 데이터 추출
                    monthly_completion_rates = step4_result.get('completion_rates', {})

                    logger.info(f"  - 3단계 법인 수: {len(logic_df)}개")
                    logger.info(f"  - 4단계 법인 수: {len(monthly_completion_rates)}개")

                    if monthly_completion_rates:
                        # 4단계 데이터를 DataFrame으로 변환
                        step4_data = []
                        for subsidiary, data_dict in monthly_completion_rates.items():
                            # completion_rates는 이제 딕셔너리 형태: {'rate': ..., 'plan_hrs': ..., 'actual_hrs': ...}
                            if isinstance(data_dict, dict):
                                rate = data_dict.get('rate', 0)
                                plan_hrs = data_dict.get('plan_hrs', 0)
                                actual_hrs = data_dict.get('actual_hrs', 0)
                            else:
                                # 만약 이전 형태(숫자)라면 호환성 유지
                                rate = data_dict
                                plan_hrs = 0
                                actual_hrs = 0

                            step4_data.append({
                                'Subsidiary': subsidiary,
                                'Planned_Hours': round(plan_hrs, 2),
                                'Actual_Hours': round(actual_hrs, 2),
                                'Hours_Completion_Rate': round(rate, 2)
                            })

                        # 4단계 DataFrame 생성
                        step4_df = pd.DataFrame(step4_data)

                        # 조인 전 법인 수 기록
                        before_join_count = len(logic_df)
                        step3_only_subsidiaries = set(logic_df['Subsidiary'].str.strip().str.upper())
                        step4_only_subsidiaries = set(step4_df['Subsidiary'].str.strip().str.upper())

                        # Outer Join 수행 (법인명으로 조인 - 대소문자 구분 없음)
                        logger.info("  - 조인 수행 중...")

                        # 대소문자 구분 없이 조인하기 위해 임시 컬럼 생성
                        logic_df['Subsidiary_lower'] = logic_df['Subsidiary'].str.lower().str.strip()
                        step4_df['Subsidiary_lower'] = step4_df['Subsidiary'].str.lower().str.strip()

                        logic_df = logic_df.merge(step4_df, left_on='Subsidiary_lower', right_on='Subsidiary_lower', how='outer', suffixes=('', '_step4'))

                        # Subsidiary 컬럼 통합 (4단계에서 온 데이터의 Subsidiary_step4를 Subsidiary로 복사)
                        logic_df['Subsidiary'] = logic_df['Subsidiary'].fillna(logic_df['Subsidiary_step4'])

                        # 임시 컬럼 제거
                        logic_df = logic_df.drop(['Subsidiary_lower', 'Subsidiary_step4'], axis=1, errors='ignore')

                        # 조인 결과 분석
                        after_join_count = len(logic_df)
                        step3_and_step4 = step3_only_subsidiaries & step4_only_subsidiaries
                        only_in_step3 = step3_only_subsidiaries - step4_only_subsidiaries
                        only_in_step4 = step4_only_subsidiaries - step3_only_subsidiaries

                        logger.info("  - 조인 결과:")
                        logger.info(f"    조인 전 (3단계): {before_join_count}개 법인")
                        logger.info(f"    조인 후: {after_join_count}개 법인")
                        logger.info(f"    양쪽 모두 존재: {len(step3_and_step4)}개 법인")
                        logger.info(f"    3단계에만 존재: {len(only_in_step3)}개 법인")
                        if only_in_step3:
                            logger.info(f"      → 과정 완료 데이터는 있지만 시간 이수율 데이터 없음")
                            logger.info(f"      → Hours 관련 컬럼이 0으로 채워짐")
                            if len(only_in_step3) <= 5:
                                logger.info(f"      법인 목록: {', '.join(sorted(only_in_step3))}")
                        logger.info(f"    4단계에만 존재: {len(only_in_step4)}개 법인")
                        if only_in_step4:
                            logger.info(f"      → 시간 이수율 데이터는 있지만 과정 완료 데이터 없음")
                            logger.info(f"      → Courses 관련 컬럼이 0으로 채워짐")
                            if len(only_in_step4) <= 5:
                                logger.info(f"      법인 목록: {', '.join(sorted(only_in_step4))}")

                        logger.info(f"✓ 4단계 데이터 조인 완료")

                        # 4단계 데이터 샘플 출력
                        logger.info("  - 4단계 데이터 샘플 (3개):")
                        for i, row in step4_df.head(3).iterrows():
                            logger.info(f"    {row['Subsidiary']}: 계획 {row['Planned_Hours']:.1f}시간, 실제 {row['Actual_Hours']:.1f}시간, 이수율 {row['Hours_Completion_Rate']:.2f}%")
                    else:
                        logger.warning("✗ 4단계 결과에서 completion_rates 데이터가 없습니다.")
                else:
                    logger.warning("✗ 4단계 결과가 없습니다.")

                # 조인 후 결측값 처리
                logger.info("8.1.2단계: 결측값을 처리합니다...")
                logger.info("  - 처리 방법:")
                logger.info("    OUTER JOIN으로 인해 발생한 null 값을 0으로 채움")
                logger.info("    대상 컬럼: Planned_Courses, Completed_Courses, Course_Completion_Rate, Planned_Hours, Actual_Hours, Hours_Completion_Rate")

                # 숫자 컬럼의 결측값을 0으로 채우기
                numeric_columns = ['Planned_Courses', 'Completed_Courses', 'Course_Completion_Rate', 'Planned_Hours', 'Actual_Hours', 'Hours_Completion_Rate']
                null_counts = {}
                for col in numeric_columns:
                    if col in logic_df.columns:
                        null_count = logic_df[col].isna().sum()
                        if null_count > 0:
                            null_counts[col] = null_count
                        logic_df[col] = logic_df[col].fillna(0)

                if null_counts:
                    logger.info("  - 컬럼별 null → 0 변환 개수:")
                    for col, count in null_counts.items():
                        logger.info(f"    {col}: {count}개")
                else:
                    logger.info("  - null 값 없음 (모든 법인이 3단계와 4단계에 모두 존재)")

                logger.info(f"✓ 결측값 처리 완료: {len(logic_df)}행")

                # 7.1.3단계: 5단계 결과와 조인
                logger.info("8.1.3단계: 5단계 결과(신입사원 교육 이수율)와 조인합니다...")
                logger.info("  - 조인 방법:")
                logger.info("    조인 타입: OUTER JOIN")
                logger.info("    조인 키: Subsidiary (대소문자 구분 없음)")
                logger.info("    Left: 기존 logic_df (3단계 + 4단계)")
                logger.info("    Right: 5단계 결과 (신입사원 교육 이수율)")
                logger.info("  - OUTER JOIN 이유:")
                logger.info("    신입사원이 없는 법인도 최종 logic.csv에 포함되어야 함")

                if step5_result is not None:
                    # 5단계 결과에서 법인별 신입사원 데이터 추출
                    subsidiary_completion = step5_result.get('subsidiary_completion', {})

                    logger.info(f"  - 기존 logic_df 법인 수: {len(logic_df)}개")
                    logger.info(f"  - 5단계 법인 수: {len(subsidiary_completion)}개")

                    if subsidiary_completion:
                        # 5단계 데이터를 DataFrame으로 변환
                        step5_data = []

                        for subsidiary, counts in subsidiary_completion.items():
                            # 이수율 = (이수 + 보류) / 전체 * 100
                            eligible_count = counts['C'] + counts['H']
                            completion_rate = (eligible_count / counts['total'] * 100) if counts['total'] > 0 else 0

                            step5_data.append({
                                'Subsidiary': subsidiary,
                                'New_Hire_Completed': counts['C'],
                                'New_Hire_Not_Completed': counts['N'],
                                'New_Hire_Pending': counts['H'],
                                'New_Hire_Total': counts['total'],
                                'New_Hire_Completion_Rate': round(completion_rate, 2)
                            })

                        # 5단계 DataFrame 생성
                        step5_df = pd.DataFrame(step5_data)

                        # 조인 전 법인 분석
                        before_join_count = len(logic_df)
                        existing_subsidiaries = set(logic_df['Subsidiary'].str.strip().str.upper())
                        step5_subsidiaries = set(step5_df['Subsidiary'].str.strip().str.upper())

                        # Outer Join 수행
                        logger.info("  - 조인 수행 중...")

                        logic_df['Subsidiary_lower'] = logic_df['Subsidiary'].str.lower().str.strip()
                        step5_df['Subsidiary_lower'] = step5_df['Subsidiary'].str.lower().str.strip()

                        logic_df = logic_df.merge(step5_df, left_on='Subsidiary_lower', right_on='Subsidiary_lower', how='outer', suffixes=('', '_step5'))

                        # Subsidiary 컬럼 통합
                        logic_df['Subsidiary'] = logic_df['Subsidiary'].fillna(logic_df['Subsidiary_step5'])

                        # 임시 컬럼 제거
                        logic_df = logic_df.drop(['Subsidiary_lower', 'Subsidiary_step5'], axis=1, errors='ignore')

                        # 조인 결과 분석
                        after_join_count = len(logic_df)
                        both_exist = existing_subsidiaries & step5_subsidiaries
                        only_in_existing = existing_subsidiaries - step5_subsidiaries
                        only_in_step5 = step5_subsidiaries - existing_subsidiaries

                        logger.info("  - 조인 결과:")
                        logger.info(f"    조인 전: {before_join_count}개 법인")
                        logger.info(f"    조인 후: {after_join_count}개 법인")
                        logger.info(f"    양쪽 모두 존재: {len(both_exist)}개 법인")
                        logger.info(f"    기존에만 존재: {len(only_in_existing)}개 법인 (신입사원 없음 → New_Hire 컬럼 0)")
                        logger.info(f"    5단계에만 존재: {len(only_in_step5)}개 법인 (신입사원만 있고 다른 데이터 없음)")

                        logger.info(f"✓ 5단계 데이터 조인 완료")

                        # 5단계 데이터 샘플 출력
                        logger.info("  - 5단계 데이터 샘플 (3개):")
                        for i, row in step5_df.head(3).iterrows():
                            logger.info(f"    {row['Subsidiary']}: 이수 {row['New_Hire_Completed']}명, 미이수 {row['New_Hire_Not_Completed']}명, 보류 {row['New_Hire_Pending']}명, 이수율 {row['New_Hire_Completion_Rate']:.2f}%")
                    else:
                        logger.warning("✗ 5단계 결과에서 subsidiary_completion 데이터가 없습니다.")
                else:
                    logger.warning("✗ 5단계 결과가 없습니다.")

                # 조인 후 결측값 재처리
                logger.info("8.1.4단계: 5단계 조인 후 결측값을 처리합니다...")
                logger.info("  - 처리 방법:")
                logger.info("    OUTER JOIN으로 인해 발생한 null 값을 0으로 채움")
                logger.info("    대상 컬럼: New_Hire 관련 5개 컬럼")

                # 숫자 컬럼의 결측값을 0으로 채우기 (5단계 컬럼 포함)
                numeric_columns = ['Planned_Courses', 'Completed_Courses', 'Course_Completion_Rate', 'Planned_Hours', 'Actual_Hours', 'Hours_Completion_Rate', 'New_Hire_Completed', 'New_Hire_Not_Completed', 'New_Hire_Pending', 'New_Hire_Total', 'New_Hire_Completion_Rate']
                null_counts = {}
                for col in numeric_columns:
                    if col in logic_df.columns:
                        null_count = logic_df[col].isna().sum()
                        if null_count > 0:
                            null_counts[col] = null_count
                        logic_df[col] = logic_df[col].fillna(0)

                if null_counts:
                    logger.info("  - 컬럼별 null → 0 변환 개수:")
                    for col, count in null_counts.items():
                        logger.info(f"    {col}: {count}개")

                logger.info(f"✓ 5단계 조인 후 결측값 처리 완료: {len(logic_df)}행")

                # 7.1.5단계: 6단계 결과와 조인
                logger.info("8.1.5단계: 6단계 결과(핵심인재 교육 이수율)와 조인합니다...")
                logger.info("  - 조인 방법:")
                logger.info("    조인 타입: OUTER JOIN")
                logger.info("    조인 키: Subsidiary (대소문자 구분 없음)")
                logger.info("    Left: 기존 logic_df (3단계 + 4단계 + 5단계)")
                logger.info("    Right: 6단계 결과 (EIP/GLP 교육 이수율)")
                logger.info("  - OUTER JOIN 이유:")
                logger.info("    핵심인재(EIP/GLP)가 없는 법인도 최종 logic.csv에 포함되어야 함")

                if step6_result is not None:
                    # 6단계 결과에서 EIP와 GLP 데이터 추출
                    eip_completion_counts = step6_result.get('eip_completion_counts', {})
                    glp_completion_counts = step6_result.get('glp_completion_counts', {})
                    total_eip = step6_result.get('total_eip', 0)
                    total_glp = step6_result.get('total_glp', 0)

                    if eip_completion_counts or glp_completion_counts:
                        # 6단계 데이터를 DataFrame으로 변환
                        step6_data = []

                        # 모든 법인 수집 (EIP와 GLP에서)
                        all_subsidiaries = set()
                        if 'eip_completion' in step6_result:
                            for emp_no in step6_result['eip_completion'].keys():
                                # 해당 직원의 법인 정보 가져오기
                                emp_data = step6_result.get('eip_employees', pd.DataFrame())
                                if not emp_data.empty and emp_no in emp_data['Employee Number'].values:
                                    subsidiary = emp_data[emp_data['Employee Number'] == emp_no]['Final Sub.'].iloc[0]
                                    all_subsidiaries.add(subsidiary)

                        if 'glp_completion' in step6_result:
                            for emp_no in step6_result['glp_completion'].keys():
                                # 해당 직원의 법인 정보 가져오기
                                emp_data = step6_result.get('glp_employees', pd.DataFrame())
                                if not emp_data.empty and emp_no in emp_data['Employee Number'].values:
                                    subsidiary = emp_data[emp_data['Employee Number'] == emp_no]['Final Sub.'].iloc[0]
                                    all_subsidiaries.add(subsidiary)

                        # 각 법인별로 EIP, GLP 통계 계산
                        for subsidiary in all_subsidiaries:
                            # EIP 통계 계산
                            eip_이수 = 0
                            eip_미이수 = 0
                            eip_전체 = 0

                            if 'eip_completion' in step6_result:
                                for emp_no, status in step6_result['eip_completion'].items():
                                    # 해당 직원의 법인 정보 확인
                                    emp_data = step6_result.get('eip_employees', pd.DataFrame())
                                    if not emp_data.empty and emp_no in emp_data['Employee Number'].values:
                                        emp_subsidiary = emp_data[emp_data['Employee Number'] == emp_no]['Final Sub.'].iloc[0]
                                        if emp_subsidiary == subsidiary:
                                            eip_전체 += 1
                                            if status == '이수':
                                                eip_이수 += 1
                                            else:
                                                eip_미이수 += 1

                            # GLP 통계 계산
                            glp_이수 = 0
                            glp_미이수 = 0
                            glp_전체 = 0

                            if 'glp_completion' in step6_result:
                                for emp_no, status in step6_result['glp_completion'].items():
                                    # 해당 직원의 법인 정보 확인
                                    emp_data = step6_result.get('glp_employees', pd.DataFrame())
                                    if not emp_data.empty and emp_no in emp_data['Employee Number'].values:
                                        emp_subsidiary = emp_data[emp_data['Employee Number'] == emp_no]['Final Sub.'].iloc[0]
                                        if emp_subsidiary == subsidiary:
                                            glp_전체 += 1
                                            if status == '이수':
                                                glp_이수 += 1
                                            else:
                                                glp_미이수 += 1

                            # 이수율 계산
                            eip_이수율 = (eip_이수 / eip_전체 * 100) if eip_전체 > 0 else 0
                            glp_이수율 = (glp_이수 / glp_전체 * 100) if glp_전체 > 0 else 0

                            step6_data.append({
                                'Subsidiary': subsidiary,
                                'EIP_Completed': eip_이수,
                                'EIP_Not_Completed': eip_미이수,
                                'EIP_Total': eip_전체,
                                'EIP_Completion_Rate': round(eip_이수율, 2),
                                'GLP_Completed': glp_이수,
                                'GLP_Not_Completed': glp_미이수,
                                'GLP_Total': glp_전체,
                                'GLP_Completion_Rate': round(glp_이수율, 2)
                            })

                        # 6단계 DataFrame 생성
                        step6_df = pd.DataFrame(step6_data)

                        logger.info(f"  - 기존 logic_df 법인 수: {len(logic_df)}개")
                        logger.info(f"  - 6단계 법인 수: {len(step6_df)}개")

                        # 조인 전 법인 분석
                        before_join_count = len(logic_df)
                        existing_subsidiaries = set(logic_df['Subsidiary'].str.strip().str.upper())
                        step6_subsidiaries = set(step6_df['Subsidiary'].str.strip().str.upper())

                        # Outer Join 수행
                        logger.info("  - 조인 수행 중...")

                        logic_df['Subsidiary_lower'] = logic_df['Subsidiary'].str.lower().str.strip()
                        step6_df['Subsidiary_lower'] = step6_df['Subsidiary'].str.lower().str.strip()

                        logic_df = logic_df.merge(step6_df, left_on='Subsidiary_lower', right_on='Subsidiary_lower', how='outer', suffixes=('', '_step6'))

                        # Subsidiary 컬럼 통합
                        logic_df['Subsidiary'] = logic_df['Subsidiary'].fillna(logic_df['Subsidiary_step6'])

                        # 임시 컬럼 제거
                        logic_df = logic_df.drop(['Subsidiary_lower', 'Subsidiary_step6'], axis=1, errors='ignore')

                        # 조인 결과 분석
                        after_join_count = len(logic_df)
                        both_exist = existing_subsidiaries & step6_subsidiaries
                        only_in_existing = existing_subsidiaries - step6_subsidiaries
                        only_in_step6 = step6_subsidiaries - existing_subsidiaries

                        logger.info("  - 조인 결과:")
                        logger.info(f"    조인 전: {before_join_count}개 법인")
                        logger.info(f"    조인 후: {after_join_count}개 법인")
                        logger.info(f"    양쪽 모두 존재: {len(both_exist)}개 법인")
                        logger.info(f"    기존에만 존재: {len(only_in_existing)}개 법인 (핵심인재 없음 → EIP/GLP 컬럼 0)")
                        logger.info(f"    6단계에만 존재: {len(only_in_step6)}개 법인 (핵심인재만 있고 다른 데이터 없음)")

                        logger.info(f"✓ 6단계 데이터 조인 완료")

                        # 6단계 데이터 샘플 출력
                        logger.info("  - 6단계 데이터 샘플 (3개):")
                        for i, row in step6_df.head(3).iterrows():
                            logger.info(f"    {row['Subsidiary']}: EIP 이수 {row['EIP_Completed']}명/{row['EIP_Total']}명({row['EIP_Completion_Rate']:.2f}%), GLP 이수 {row['GLP_Completed']}명/{row['GLP_Total']}명({row['GLP_Completion_Rate']:.2f}%)")
                    else:
                        logger.warning("✗ 6단계 결과에서 EIP/GLP 데이터가 없습니다.")
                else:
                    logger.warning("✗ 6단계 결과가 없습니다.")

                # 조인 후 결측값 최종 처리
                logger.info("8.1.6단계: 6단계 조인 후 결측값을 최종 처리합니다...")
                logger.info("  - 처리 방법:")
                logger.info("    OUTER JOIN으로 인해 발생한 null 값을 0으로 채움")
                logger.info("    대상 컬럼: EIP/GLP 관련 8개 컬럼")

                # 숫자 컬럼의 결측값을 0으로 채우기 (6단계 컬럼 포함)
                numeric_columns = ['Planned_Courses', 'Completed_Courses', 'Course_Completion_Rate', 'Planned_Hours', 'Actual_Hours', 'Hours_Completion_Rate', 'New_Hire_Completed', 'New_Hire_Not_Completed', 'New_Hire_Pending', 'New_Hire_Total', 'New_Hire_Completion_Rate', 'EIP_Completed', 'EIP_Not_Completed', 'EIP_Total', 'EIP_Completion_Rate', 'GLP_Completed', 'GLP_Not_Completed', 'GLP_Total', 'GLP_Completion_Rate']
                null_counts = {}
                for col in numeric_columns:
                    if col in logic_df.columns:
                        null_count = logic_df[col].isna().sum()
                        if null_count > 0:
                            null_counts[col] = null_count
                        logic_df[col] = logic_df[col].fillna(0)

                if null_counts:
                    logger.info("  - 컬럼별 null → 0 변환 개수:")
                    for col, count in null_counts.items():
                        logger.info(f"    {col}: {count}개")

                logger.info(f"✓ 6단계 조인 후 결측값 최종 처리 완료: {len(logic_df)}행")

                # 7.1.7단계: 7단계 결과와 조인 (신입 팀장)
                logger.info("8.1.7단계: 7단계 결과(신입 팀장 교육 이수율)와 조인합니다...")
                logger.info("  - 조인 방법:")
                logger.info("    조인 타입: OUTER JOIN")
                logger.info("    조인 키: Subsidiary (대소문자 구분 없음)")
                logger.info("    Left: 기존 logic_df (3단계 + 4단계 + 5단계 + 6단계)")
                logger.info("    Right: 7단계 결과 (신입 팀장 교육 이수율)")
                logger.info("  - OUTER JOIN 이유:")
                logger.info("    신입 팀장이 없는 법인도 최종 logic.csv에 포함되어야 함")

                if step7_result is not None:
                    # 7단계 결과에서 신입 팀장 데이터 추출
                    new_leader_completion_counts = step7_result.get('new_leader_completion_counts', {})
                    total_new_leader = step7_result.get('total_new_leader', 0)

                    if new_leader_completion_counts or total_new_leader > 0:
                        # 7단계 데이터를 DataFrame으로 변환
                        step7_data = []

                        # 모든 법인 수집
                        all_subsidiaries = set()
                        if 'new_leader_completion' in step7_result:
                            for emp_no in step7_result['new_leader_completion'].keys():
                                # 해당 직원의 법인 정보 가져오기
                                emp_data = step7_result.get('new_leader_employees', pd.DataFrame())
                                if not emp_data.empty and emp_no in emp_data['Employee Number'].values:
                                    subsidiary = emp_data[emp_data['Employee Number'] == emp_no]['Final Sub.'].iloc[0]
                                    all_subsidiaries.add(subsidiary)

                        # 각 법인별로 신입 팀장 통계 계산
                        for subsidiary in all_subsidiaries:
                            # 신입 팀장 통계 계산
                            new_leader_이수 = 0
                            new_leader_미이수 = 0
                            new_leader_전체 = 0

                            if 'new_leader_completion' in step7_result:
                                for emp_no, status in step7_result['new_leader_completion'].items():
                                    # 해당 직원의 법인 정보 확인
                                    emp_data = step7_result.get('new_leader_employees', pd.DataFrame())
                                    if not emp_data.empty and emp_no in emp_data['Employee Number'].values:
                                        emp_subsidiary = emp_data[emp_data['Employee Number'] == emp_no]['Final Sub.'].iloc[0]
                                        if emp_subsidiary == subsidiary:
                                            new_leader_전체 += 1
                                            if status == '이수':
                                                new_leader_이수 += 1
                                            elif status == '미이수':
                                                new_leader_미이수 += 1

                            # 신입 팀장 이수율 계산
                            new_leader_이수율 = (new_leader_이수 / new_leader_전체 * 100) if new_leader_전체 > 0 else 0

                            step7_data.append({
                                'Subsidiary': subsidiary,
                                'New_Leader_Completed': new_leader_이수,
                                'New_Leader_Not_Completed': new_leader_미이수,
                                'New_Leader_Total': new_leader_전체,
                                'New_Leader_Completion_Rate': round(new_leader_이수율, 2)
                            })

                        # DataFrame 생성
                        step7_df = pd.DataFrame(step7_data)

                        logger.info(f"  - 7단계에서 {len(step7_df)}개 법인의 신입 팀장 교육 데이터 생성")

                        # 조인 전 법인 수 저장
                        before_join_count = len(logic_df)
                        existing_subsidiaries = set(logic_df['Subsidiary'].str.lower().str.strip().unique())
                        step7_subsidiaries = set(step7_df['Subsidiary'].str.lower().str.strip().unique())

                        logger.info("  - 조인 수행 중...")

                        logic_df['Subsidiary_lower'] = logic_df['Subsidiary'].str.lower().str.strip()
                        step7_df['Subsidiary_lower'] = step7_df['Subsidiary'].str.lower().str.strip()

                        logic_df = logic_df.merge(step7_df, left_on='Subsidiary_lower', right_on='Subsidiary_lower', how='outer', suffixes=('', '_step7'))

                        # Subsidiary 컬럼 통합
                        logic_df['Subsidiary'] = logic_df['Subsidiary'].fillna(logic_df['Subsidiary_step7'])

                        # 임시 컬럼 제거
                        logic_df = logic_df.drop(['Subsidiary_lower', 'Subsidiary_step7'], axis=1, errors='ignore')

                        # 조인 결과 분석
                        after_join_count = len(logic_df)
                        both_exist = existing_subsidiaries & step7_subsidiaries
                        only_in_existing = existing_subsidiaries - step7_subsidiaries
                        only_in_step7 = step7_subsidiaries - existing_subsidiaries

                        logger.info("  - 조인 결과:")
                        logger.info(f"    조인 전: {before_join_count}개 법인")
                        logger.info(f"    조인 후: {after_join_count}개 법인")
                        logger.info(f"    양쪽 모두 존재: {len(both_exist)}개 법인")
                        logger.info(f"    기존에만 존재: {len(only_in_existing)}개 법인 (신입 팀장 없음 → New_Leader 컬럼 0)")
                        logger.info(f"    7단계에만 존재: {len(only_in_step7)}개 법인 (신입 팀장만 있고 다른 데이터 없음)")

                        logger.info(f"✓ 7단계 데이터 조인 완료")

                        # 7단계 데이터 샘플 출력
                        logger.info("  - 7단계 데이터 샘플 (3개):")
                        for i, row in step7_df.head(3).iterrows():
                            logger.info(f"    {row['Subsidiary']}: 신입 팀장 이수 {row['New_Leader_Completed']}명/{row['New_Leader_Total']}명({row['New_Leader_Completion_Rate']:.2f}%)")
                    else:
                        logger.warning("✗ 7단계 결과에서 신입 팀장 데이터가 없습니다.")
                else:
                    logger.warning("✗ 7단계 결과가 없습니다.")

                # 조인 후 결측값 최종 처리 (신입 팀장 포함)
                logger.info("8.1.8단계: 7단계 조인 후 결측값을 최종 처리합니다...")
                logger.info("  - 처리 방법:")
                logger.info("    OUTER JOIN으로 인해 발생한 null 값을 0으로 채움")
                logger.info("    대상 컬럼: 신입 팀장 관련 4개 컬럼")

                # 숫자 컬럼의 결측값을 0으로 채우기 (7단계 컬럼 포함)
                numeric_columns = ['Planned_Courses', 'Completed_Courses', 'Course_Completion_Rate', 'Planned_Hours', 'Actual_Hours', 'Hours_Completion_Rate', 'New_Hire_Completed', 'New_Hire_Not_Completed', 'New_Hire_Pending', 'New_Hire_Total', 'New_Hire_Completion_Rate', 'EIP_Completed', 'EIP_Not_Completed', 'EIP_Total', 'EIP_Completion_Rate', 'GLP_Completed', 'GLP_Not_Completed', 'GLP_Total', 'GLP_Completion_Rate', 'New_Leader_Completed', 'New_Leader_Not_Completed', 'New_Leader_Total', 'New_Leader_Completion_Rate']
                null_counts = {}
                for col in numeric_columns:
                    if col in logic_df.columns:
                        null_count = logic_df[col].isna().sum()
                        if null_count > 0:
                            null_counts[col] = null_count
                        logic_df[col] = logic_df[col].fillna(0)

                if null_counts:
                    logger.info("  - 컬럼별 null → 0 변환 개수:")
                    for col, count in null_counts.items():
                        logger.info(f"    {col}: {count}개")

                logger.info(f"✓ 7단계 조인 후 결측값 최종 처리 완료: {len(logic_df)}행")

                # 8.3단계: hr_index_final과 조인하여 Final Region 추가
                logger.info("8.3단계: hr_index_final.csv와 조인하여 Final Region 컬럼을 추가합니다...")
                logger.info("  - 조인 방법:")
                logger.info("    조인 타입: LEFT JOIN")
                logger.info("    조인 키: Subsidiary = Final Sub. (대소문자 구분 없음)")
                logger.info("    Left: 기존 logic_df (3~7단계 모든 데이터)")
                logger.info("    Right: hr_index_final.csv (Final Region 정보)")
                logger.info("  - LEFT JOIN 이유:")
                logger.info("    logic_df의 모든 법인을 유지하면서 Final Region 정보만 추가")
                logger.info("    매칭되지 않으면 Final Region은 null")

                try:
                    # hr_index_final.csv 파일 로드
                    hr_index_path = os.path.join(file_directory, "hr_index_final.csv")

                    if os.path.exists(hr_index_path):
                        logger.info(f"✓ hr_index_final.csv 파일을 로드합니다: {hr_index_path}")
                        hr_index_df = pd.read_csv(hr_index_path, encoding='utf-8-sig')

                        # Final Sub.와 Final Region 컬럼이 있는지 확인
                        if 'Final Sub.' in hr_index_df.columns and 'Final Region' in hr_index_df.columns:
                            logger.info(f"✓ hr_index_final.csv 컬럼 확인: {list(hr_index_df.columns)}")

                            # 중복 제거하여 각 법인별 Final Region 매핑 생성
                            region_mapping = hr_index_df[['Final Sub.', 'Final Region']].drop_duplicates()

                            logger.info(f"  - hr_index_final.csv 법인 수: {len(region_mapping)}개")
                            logger.info(f"  - 기존 logic_df 법인 수: {len(logic_df)}개")

                            # logic_df와 조인 (Left Join - 대소문자 구분 없음)
                            logic_df_before = len(logic_df)

                            # 조인 전 법인 분석
                            logic_subsidiaries = set(logic_df['Subsidiary'].str.strip().str.upper())
                            hr_subsidiaries = set(region_mapping['Final Sub.'].str.strip().str.upper())

                            logger.info("  - 조인 수행 중...")

                            # 대소문자 구분 없이 조인하기 위해 임시 컬럼 생성
                            logic_df['Subsidiary_lower'] = logic_df['Subsidiary'].str.lower().str.strip()
                            region_mapping['Final_Sub_lower'] = region_mapping['Final Sub.'].str.lower().str.strip()

                            logic_df = logic_df.merge(region_mapping, left_on='Subsidiary_lower', right_on='Final_Sub_lower', how='left', suffixes=('', '_region'))

                            # 임시 컬럼 제거
                            logic_df = logic_df.drop(['Subsidiary_lower', 'Final_Sub_lower', 'Final Sub._region'], axis=1, errors='ignore')

                            # 불필요한 컬럼 제거 (Final Sub. 중복)
                            if 'Final Sub.' in logic_df.columns:
                                logic_df = logic_df.drop('Final Sub.', axis=1)

                            # 조인 결과 분석
                            matched_count = logic_df['Final Region'].notna().sum()
                            unmatched_count = logic_df['Final Region'].isna().sum()
                            both_exist = logic_subsidiaries & hr_subsidiaries
                            only_in_logic = logic_subsidiaries - hr_subsidiaries

                            logger.info("  - 조인 결과:")
                            logger.info(f"    조인 전: {logic_df_before}개 법인")
                            logger.info(f"    조인 후: {len(logic_df)}개 법인 (LEFT JOIN이므로 동일)")
                            logger.info(f"    Final Region 매칭 성공: {matched_count}개 법인")
                            logger.info(f"    Final Region 매칭 실패: {unmatched_count}개 법인")

                            if unmatched_count > 0:
                                logger.warning(f"      → hr_index_final.csv에 없는 법인: {unmatched_count}개")
                                unmatched_subsidiaries = logic_df[logic_df['Final Region'].isna()]['Subsidiary'].tolist()
                                if len(unmatched_subsidiaries) <= 10:
                                    logger.warning(f"      법인 목록: {', '.join(unmatched_subsidiaries)}")
                                else:
                                    logger.warning(f"      법인 목록 (처음 10개): {', '.join(unmatched_subsidiaries[:10])}...")

                            logger.info(f"✓ Final Region 조인 완료")

                            # Final Region 샘플 데이터 출력
                            logger.info("  - Final Region 매핑 샘플 (5개):")
                            sample_data = logic_df[['Subsidiary', 'Final Region']].head(5)
                            for i, row in sample_data.iterrows():
                                region = row['Final Region'] if pd.notna(row['Final Region']) else 'null'
                                logger.info(f"    {row['Subsidiary']}: {region}")

                        else:
                            logger.warning("✗ hr_index_final.csv에서 'Final Sub.' 또는 'Final Region' 컬럼을 찾을 수 없습니다.")
                            logger.info(f"✓ 사용 가능한 컬럼: {list(hr_index_df.columns)}")
                    else:
                        logger.warning(f"✗ hr_index_final.csv 파일을 찾을 수 없습니다: {hr_index_path}")

                except Exception as e:
                    logger.error(f"✗ Final Region 조인 중 오류 발생: {e}")

                # 7.1.4단계: Completion Rate 100% 제한
                logger.info("8.1.4단계: Completion Rate 컬럼 100% 제한")
                rate_columns = ['Course_Completion_Rate', 'Hours_Completion_Rate', 'New_Hire_Completion_Rate',
                              'EIP_Completion_Rate', 'GLP_Completion_Rate']

                over_100_counts = {}
                for col in rate_columns:
                    if col in logic_df.columns:
                        # 100을 초과하는 값 찾기
                        over_100_mask = logic_df[col] > 100
                        over_100_count = over_100_mask.sum()

                        if over_100_count > 0:
                            over_100_counts[col] = over_100_count
                            # 100 초과하는 법인 출력
                            over_100_subs = logic_df.loc[over_100_mask, ['Subsidiary', col]]
                            logger.info(f"  - {col}: {over_100_count}개 법인이 100% 초과")
                            for idx, row in over_100_subs.iterrows():
                                logger.info(f"    {row['Subsidiary']}: {row[col]:.2f}% → 100.0%")

                            # 100으로 제한
                            logic_df.loc[over_100_mask, col] = 100.0

                if over_100_counts:
                    logger.info(f"✓ 7.1.4단계 완료: {sum(over_100_counts.values())}건의 100% 초과 값을 100%로 제한")
                else:
                    logger.info(f"✓ 7.1.4단계 완료: 100% 초과 값 없음")

                # CSV 파일로 저장
                output_path = os.path.join(file_directory, "logic.csv")
                logic_df.to_csv(output_path, index=False, encoding='utf-8-sig')

                logger.info(f"✓ logic.csv 파일 저장 완료: {output_path}")
                logger.info(f"✓ 저장된 데이터: {len(logic_df)}행, {len(logic_df.columns)}열")

                # 7.1.7단계: Index Management와 비교
                logger.info("8.1.7단계: Index Management와 법인 비교...")
                logger.info("  - 목적: Manage Area = 'Y'인 법인과 logic.csv 법인 비교")

                try:
                    index_mgmt_path = os.path.join(file_directory, "index_management_final.csv")
                    if os.path.exists(index_mgmt_path):
                        index_mgmt_df = pd.read_csv(index_mgmt_path, encoding='utf-8-sig')

                        if 'Final Sub.' in index_mgmt_df.columns and 'Manage Area' in index_mgmt_df.columns:
                            # Manage Area = 'Y'인 법인만 추출
                            mgmt_y_subs = set(index_mgmt_df[index_mgmt_df['Manage Area'] == 'Y']['Final Sub.'].str.strip().str.upper())
                            logic_subs = set(logic_df['Subsidiary'].str.strip().str.upper())

                            logger.info(f"  - Index Management (Manage Area = Y): {len(mgmt_y_subs)}개 법인")
                            logger.info(f"  - Logic.csv: {len(logic_subs)}개 법인")

                            # Index Management에는 있지만 Logic에는 없는 법인
                            only_in_mgmt = mgmt_y_subs - logic_subs
                            if only_in_mgmt:
                                logger.warning(f"  - Index Management에만 있음 (Logic에 누락): {len(only_in_mgmt)}개")
                                logger.warning(f"    법인 목록: {', '.join(sorted(only_in_mgmt))}")
                                logger.warning("    → 이 법인들은 3~6단계 어디에도 데이터가 없어서 logic.csv에 포함되지 않음")
                                logger.warning("    → hong_data_plan_final.csv에 교육 계획이 있는지 확인 필요")
                            else:
                                logger.info("  - Index Management에만 있는 법인: 없음")

                            # Logic에는 있지만 Index Management에는 없는 법인
                            only_in_logic = logic_subs - mgmt_y_subs
                            if only_in_logic:
                                logger.warning(f"  - Logic에만 있음 (Index Management에 없음): {len(only_in_logic)}개")
                                logger.warning(f"    법인 목록: {', '.join(sorted(only_in_logic))}")
                                logger.warning("    → 데이터는 있지만 관리 대상(Manage Area=Y)이 아닌 법인")
                            else:
                                logger.info("  - Logic에만 있는 법인: 없음")

                            # 양쪽 모두 존재
                            both = mgmt_y_subs & logic_subs
                            logger.info(f"  - 양쪽 모두 존재: {len(both)}개 법인")
                            logger.info(f"✓ 7.1.7단계 완료: Index Management 비교 완료")
                        else:
                            logger.warning("  - Index Management에 필요한 컬럼이 없어 비교를 건너뜁니다.")
                    else:
                        logger.warning(f"  - Index Management 파일을 찾을 수 없어 비교를 건너뜁니다: {index_mgmt_path}")
                except Exception as e:
                    logger.warning(f"  - Index Management 비교 중 오류 발생: {e}")

                # 7.1.8단계: Index Management와 조인하여 추가 정보 병합
                logger.info("8.1.8단계: Index Management와 조인하여 추가 정보를 병합합니다...")
                logger.info("  - 조인 방법:")
                logger.info("    조인 타입: LEFT JOIN")
                logger.info("    조인 키: Subsidiary = Final Sub. (대소문자 구분 없음)")
                logger.info("    Left: logic_df")
                logger.info("    Right: index_management_final.csv")
                logger.info("  - 제외 컬럼: Final Region (이미 존재), Manage Area (필터링용)")
                logger.info("  - 목적: Index Management의 추가 정보를 logic.csv에 병합")

                try:
                    index_mgmt_path = os.path.join(file_directory, "index_management_final.csv")
                    if os.path.exists(index_mgmt_path):
                        index_mgmt_df = pd.read_csv(index_mgmt_path, encoding='utf-8-sig')

                        if 'Final Sub.' in index_mgmt_df.columns:
                            logger.info(f"  - Index Management 컬럼 수: {len(index_mgmt_df.columns)}개")
                            logger.info(f"  - Index Management 법인 수: {len(index_mgmt_df)}개")

                            # 제외할 컬럼 목록
                            exclude_columns = ['Final Region', 'Manage Area']

                            # 병합할 컬럼 선택 (Final Sub. + 제외 컬럼 제외)
                            merge_columns = [col for col in index_mgmt_df.columns if col not in exclude_columns]
                            logger.info(f"  - 병합 대상 컬럼 ({len(merge_columns)}개):")
                            for col in merge_columns:
                                logger.info(f"    {col}")

                            # 조인 전 컬럼 수
                            before_columns = len(logic_df.columns)

                            # LEFT JOIN 수행
                            logger.info("  - LEFT JOIN 수행 중...")
                            logic_df['Subsidiary_upper'] = logic_df['Subsidiary'].str.strip().str.upper()
                            index_mgmt_df['Final_Sub_upper'] = index_mgmt_df['Final Sub.'].str.strip().str.upper()

                            # 병합할 컬럼만 선택
                            index_mgmt_merge = index_mgmt_df[['Final_Sub_upper'] + [col for col in merge_columns if col != 'Final Sub.']].copy()

                            logic_df = logic_df.merge(
                                index_mgmt_merge,
                                left_on='Subsidiary_upper',
                                right_on='Final_Sub_upper',
                                how='left',
                                suffixes=('', '_index')
                            )

                            # 임시 컬럼 제거
                            logic_df = logic_df.drop(['Subsidiary_upper', 'Final_Sub_upper'], axis=1, errors='ignore')

                            # 조인 후 컬럼 수
                            after_columns = len(logic_df.columns)
                            added_columns = after_columns - before_columns

                            logger.info(f"  - 조인 결과:")
                            logger.info(f"    조인 전 컬럼 수: {before_columns}개")
                            logger.info(f"    조인 후 컬럼 수: {after_columns}개")
                            logger.info(f"    추가된 컬럼 수: {added_columns}개")

                            # 매칭 통계
                            # 첫 번째 병합된 컬럼의 null 개수로 매칭 실패 건수 추정
                            first_merged_col = [col for col in merge_columns if col != 'Final Sub.'][0] if len(merge_columns) > 1 else None
                            if first_merged_col and first_merged_col in logic_df.columns:
                                matched = logic_df[first_merged_col].notna().sum()
                                unmatched = logic_df[first_merged_col].isna().sum()
                                logger.info(f"    매칭 성공: {matched}개 법인")
                                logger.info(f"    매칭 실패: {unmatched}개 법인")

                            logger.info(f"✓ 7.1.8단계 완료: Index Management 조인 완료")

                            # 7.1.9단계: Completion Rate 100% 제한
                            logger.info("8.1.9단계: Completion Rate 컬럼 100% 제한")
                            rate_columns = ['Course_Completion_Rate', 'Hours_Completion_Rate', 'New_Hire_Completion_Rate',
                                          'EIP_Completion_Rate', 'GLP_Completion_Rate']

                            over_100_counts = {}
                            for col in rate_columns:
                                if col in logic_df.columns:
                                    # 100을 초과하는 값 찾기
                                    over_100_mask = logic_df[col] > 100
                                    over_100_count = over_100_mask.sum()

                                    if over_100_count > 0:
                                        over_100_counts[col] = over_100_count
                                        # 100 초과하는 법인 출력
                                        over_100_subs = logic_df.loc[over_100_mask, ['Subsidiary', col]]
                                        logger.info(f"  - {col}: {over_100_count}개 법인이 100% 초과")
                                        for idx, row in over_100_subs.iterrows():
                                            logger.info(f"    {row['Subsidiary']}: {row[col]:.2f}% → 100.0%")

                                        # 100으로 제한
                                        logic_df.loc[over_100_mask, col] = 100.0

                            if over_100_counts:
                                logger.info(f"✓ 7.1.9단계 완료: {sum(over_100_counts.values())}건의 100% 초과 값을 100%로 제한")
                            else:
                                logger.info(f"✓ 7.1.9단계 완료: 100% 초과 값 없음")

                            # 7.3단계: Score 계산
                            logger.info("")
                            logger.info("8.3.1단계: Score 계산 (Monthly Index + Quarterly Index = 100점)")
                            logger.info("=" * 80)

                            # 분기 결정
                            if ANALYSIS_MONTH in [1, 2, 3]:
                                quarter = 1
                            elif ANALYSIS_MONTH in [4, 5, 6]:
                                quarter = 2
                            elif ANALYSIS_MONTH in [7, 8, 9]:
                                quarter = 3
                            else:  # 10, 11, 12
                                quarter = 4

                            logger.info(f"  - 분석 월: {ANALYSIS_MONTH}월")
                            logger.info(f"  - 분석 분기: {quarter}Q")
                            logger.info("")

                            # Score 계산 함수
                            def calculate_rate_score(rate, total, score_table, check_total=True):
                                """
                                완료율 기반 점수 계산
                                rate: 완료율 (%)
                                total: 전체 인원/과정 수
                                score_table: {min_rate: score} 딕셔너리 (내림차순 정렬 필요)
                                check_total: Total이 0일 때 만점 처리 여부 (기본 True)
                                """
                                # Total 체크 (New Hire, EIP, GLP만 해당)
                                if check_total and (pd.isna(total) or total == 0):
                                    return max(score_table.values())

                                # Rate가 NaN이면 최소점
                                if pd.isna(rate):
                                    return min(score_table.values())

                                # Rate 기준으로 점수 결정
                                for min_rate, score in sorted(score_table.items(), reverse=True):
                                    if rate >= min_rate:
                                        return score

                                return min(score_table.values())

                            def calculate_yn_score(yn_value, max_score):
                                """Y/N 기반 점수 계산"""
                                if pd.isna(yn_value):
                                    return 0
                                return max_score if str(yn_value).strip().upper() == 'Y' else 0

                            # 점수 기준표
                            logger.info("7.3.1: 점수 기준표")
                            logger.info("  - LMS Course Registration Rate (10점): 90%≥10, 80-89%=8, 60-79%=6, 40-59%=4, <40%=2")
                            logger.info("  - Plan vs Execution Rate (20점): 90%≥20, 80-89%=16, 60-79%=12, 40-59%=8, <40%=4")
                            logger.info("  - New Hire Completion Rate (15점): Total=0→15, 90%≥15, 80-89%=12, 60-79%=10, 40-59%=8, <40%=6")
                            logger.info("  - HIPO(EIP) Completion Rate (15점): Total=0→15, 90%≥15, 80-89%=12, 60-79%=10, 40-59%=8, <40%=6")
                            logger.info("  - HIPO(GLP) Completion Rate (15점): Total=0→15, 90%≥15, 80-89%=12, 60-79%=10, 40-59%=8, <40%=6")
                            logger.info("  - New Leader Completion Rate (15점): Total=0→15, 90%≥15, 80-89%=12, 60-79%=10, 40-59%=8, <40%=6")
                            logger.info("  - Y/N 지표: Y=만점, N=0점")
                            logger.info("")

                            # Monthly Index 계산 (70점)
                            logger.info("7.3.2: Monthly Index 계산 (70점)")

                            # LMS Course Registration Rate (10점) - Total 체크 안함
                            lms_course_scores = logic_df.apply(
                                lambda row: calculate_rate_score(
                                    row.get('Course_Completion_Rate'),
                                    row.get('Planned_Courses'),
                                    {90: 10, 80: 8, 60: 6, 40: 4, 0: 2},
                                    check_total=False
                                ), axis=1
                            )
                            logger.info(f"  - LMS Course Registration Rate: 평균 {lms_course_scores.mean():.2f}점")

                            # Plan vs Execution Rate (20점) - Total 체크 안함
                            plan_exec_scores = logic_df.apply(
                                lambda row: calculate_rate_score(
                                    row.get('Hours_Completion_Rate'),
                                    row.get('Planned_Hours'),
                                    {90: 20, 80: 16, 60: 12, 40: 8, 0: 4},
                                    check_total=False
                                ), axis=1
                            )
                            logger.info(f"  - Plan vs Execution Rate: 평균 {plan_exec_scores.mean():.2f}점")

                            # New Hire Completion Rate (15점)
                            new_hire_scores = logic_df.apply(
                                lambda row: calculate_rate_score(
                                    row.get('New_Hire_Completion_Rate'),
                                    row.get('New_Hire_Total'),
                                    {90: 15, 80: 12, 60: 10, 40: 8, 0: 6}
                                ), axis=1
                            )
                            logger.info(f"  - New Hire Completion Rate: 평균 {new_hire_scores.mean():.2f}점")

                            # New Leader Completion Rate (15점)
                            new_leader_scores = logic_df.apply(
                                lambda row: calculate_rate_score(
                                    row.get('New_Leader_Completion_Rate'),
                                    row.get('New_Leader_Total'),
                                    {90: 15, 80: 12, 60: 10, 40: 8, 0: 6}
                                ), axis=1
                            )
                            logger.info(f"  - New Leader Completion Rate: 평균 {new_leader_scores.mean():.2f}점")

                            # JAM Member (10점)
                            jam_scores = logic_df.apply(
                                lambda row: calculate_yn_score(row.get('JAM Member'), 10), axis=1
                            )
                            logger.info(f"  - JAM Member: 평균 {jam_scores.mean():.2f}점")

                            monthly_index = lms_course_scores + plan_exec_scores + new_hire_scores + new_leader_scores + jam_scores
                            logger.info(f"  ✓ Monthly Index 평균: {monthly_index.mean():.2f}점 / 70점")
                            logger.info("")

                            # Quarterly Index 계산 (30점)
                            logger.info(f"7.3.3: Quarterly Index 계산 (30점) - {quarter}Q 기준")

                            if quarter == 1:
                                # 1Q: New LMS Course (10) + LMS Mission (10) + Annual Plan Setup (10)
                                logger.info("  - 1Q 적용 지표: New LMS Course, LMS Mission, Annual Plan Setup")
                                new_lms_scores = logic_df.apply(
                                    lambda row: calculate_yn_score(row.get('New LMS Course'), 10), axis=1
                                )
                                lms_mission_scores = logic_df.apply(
                                    lambda row: calculate_yn_score(row.get('LMS Mission'), 10), axis=1
                                )
                                annual_plan_scores = logic_df.apply(
                                    lambda row: calculate_yn_score(row.get('Annual Plan Setup'), 10), axis=1
                                )
                                quarterly_index = new_lms_scores + lms_mission_scores + annual_plan_scores
                                logger.info(f"    New LMS Course: 평균 {new_lms_scores.mean():.2f}점")
                                logger.info(f"    LMS Mission: 평균 {lms_mission_scores.mean():.2f}점")
                                logger.info(f"    Annual Plan Setup: 평균 {annual_plan_scores.mean():.2f}점")

                            elif quarter in [2, 3]:
                                # 2Q, 3Q: HIPO(EIP) (15) + HIPO(GLP) (15)
                                logger.info("  - 2Q/3Q 적용 지표: HIPO(EIP) Completion Rate, HIPO(GLP) Completion Rate")
                                eip_scores = logic_df.apply(
                                    lambda row: calculate_rate_score(
                                        row.get('EIP_Completion_Rate'),
                                        row.get('EIP_Total'),
                                        {90: 15, 80: 12, 60: 10, 40: 8, 0: 6}
                                    ), axis=1
                                )
                                glp_scores = logic_df.apply(
                                    lambda row: calculate_rate_score(
                                        row.get('GLP_Completion_Rate'),
                                        row.get('GLP_Total'),
                                        {90: 15, 80: 12, 60: 10, 40: 8, 0: 6}
                                    ), axis=1
                                )
                                quarterly_index = eip_scores + glp_scores
                                logger.info(f"    HIPO(EIP) Completion Rate: 평균 {eip_scores.mean():.2f}점")
                                logger.info(f"    HIPO(GLP) Completion Rate: 평균 {glp_scores.mean():.2f}점")

                            else:  # quarter == 4
                                # 4Q: Global L&D Council (15) + Infra index response (15)
                                logger.info("  - 4Q 적용 지표: Global L&D Council, Infra index response")
                                global_ld_scores = logic_df.apply(
                                    lambda row: calculate_yn_score(row.get('Global L&D Council'), 15), axis=1
                                )
                                infra_scores = logic_df.apply(
                                    lambda row: calculate_yn_score(row.get('Infra index response'), 15), axis=1
                                )
                                quarterly_index = global_ld_scores + infra_scores
                                logger.info(f"    Global L&D Council: 평균 {global_ld_scores.mean():.2f}점")
                                logger.info(f"    Infra index response: 평균 {infra_scores.mean():.2f}점")

                            logger.info(f"  ✓ Quarterly Index 평균: {quarterly_index.mean():.2f}점 / 30점")
                            logger.info("")

                            # 총점 계산
                            logger.info("7.3.4: 총점 계산 (Monthly + Quarterly)")
                            logic_df['Score'] = monthly_index + quarterly_index
                            logger.info(f"  ✓ Score 컬럼 추가 완료")
                            logger.info(f"  ✓ Score 평균: {logic_df['Score'].mean():.2f}점 / 100점")
                            logger.info(f"  ✓ Score 최대: {logic_df['Score'].max():.2f}점")
                            logger.info(f"  ✓ Score 최소: {logic_df['Score'].min():.2f}점")

                            # Score 분포 출력 (상위 5개, 하위 5개)
                            logger.info("")
                            logger.info("  - Score 상위 5개 법인:")
                            top_5 = logic_df.nlargest(5, 'Score')[['Subsidiary', 'Score']]
                            for idx, row in top_5.iterrows():
                                logger.info(f"    {row['Subsidiary']}: {row['Score']:.2f}점")

                            logger.info("")
                            logger.info("  - Score 하위 5개 법인:")
                            bottom_5 = logic_df.nsmallest(5, 'Score')[['Subsidiary', 'Score']]
                            for idx, row in bottom_5.iterrows():
                                logger.info(f"    {row['Subsidiary']}: {row['Score']:.2f}점")

                            logger.info("")
                            logger.info(f"✓ 7.3단계 완료: Score 계산 완료")
                            logger.info("=" * 80)
                            logger.info("")

                            # 최종 logic.csv 다시 저장
                            logger.info("  - 최종 logic.csv 다시 저장 중...")
                            logic_df.to_csv(output_path, index=False, encoding='utf-8-sig')
                            logger.info(f"    ✓ 업데이트된 logic.csv 저장 완료: {output_path}")
                            logger.info(f"    ✓ 최종 데이터: {len(logic_df)}행, {len(logic_df.columns)}열")
                        else:
                            logger.warning("  - Index Management에 'Final Sub.' 컬럼이 없어 조인을 건너뜁니다.")
                    else:
                        logger.warning(f"  - Index Management 파일을 찾을 수 없어 조인을 건너뜁니다: {index_mgmt_path}")
                except Exception as e:
                    logger.warning(f"  - Index Management 조인 중 오류 발생: {e}")
                    import traceback
                    logger.warning(f"  - 상세 오류: {traceback.format_exc()}")

                # 샘플 데이터 출력
                logger.info("최종 logic.csv 샘플 데이터:")
                for i, row in logic_df.head(3).iterrows():
                    # Final Region 정보 추가
                    region_info = f", 지역: {row.get('Final Region', '없음')}" if 'Final Region' in logic_df.columns else ""

                    # 6단계 컬럼이 있는지 확인
                    if 'EIP_Completed' in logic_df.columns:
                        logger.info(f"  {row['Subsidiary']}: 진행중 {row['Planned_Courses']}개, 완료 {row['Completed_Courses']}개, 완료율 {row['Course_Completion_Rate']:.2f}%, 계획시간 {row['Planned_Hours']:.1f}시간, 수강시간 {row['Actual_Hours']:.1f}시간, 이수율 {row['Hours_Completion_Rate']:.2f}%, 신입이수 {row['New_Hire_Completed']}명, 신입미이수 {row['New_Hire_Not_Completed']}명, 신입보류 {row['New_Hire_Pending']}명, 신입전체 {row['New_Hire_Total']}명, 신입이수율 {row['New_Hire_Completion_Rate']:.2f}%, EIP이수 {row['EIP_Completed']}명, EIP미이수 {row['EIP_Not_Completed']}명, EIP전체 {row['EIP_Total']}명, EIP이수율 {row['EIP_Completion_Rate']:.2f}%, GLP이수 {row['GLP_Completed']}명, GLP미이수 {row['GLP_Not_Completed']}명, GLP전체 {row['GLP_Total']}명, GLP이수율 {row['GLP_Completion_Rate']:.2f}%{region_info}")
                    elif 'New_Hire_Completed' in logic_df.columns:
                        logger.info(f"  {row['Subsidiary']}: 진행중 {row['Planned_Courses']}개, 완료 {row['Completed_Courses']}개, 완료율 {row['Course_Completion_Rate']:.2f}%, 계획시간 {row['Planned_Hours']:.1f}시간, 수강시간 {row['Actual_Hours']:.1f}시간, 이수율 {row['Hours_Completion_Rate']:.2f}%, 신입이수 {row['New_Hire_Completed']}명, 신입미이수 {row['New_Hire_Not_Completed']}명, 신입보류 {row['New_Hire_Pending']}명, 신입전체 {row['New_Hire_Total']}명, 신입이수율 {row['New_Hire_Completion_Rate']:.2f}%{region_info}")
                    elif 'Planned_Hours' in logic_df.columns:
                        logger.info(f"  {row['Subsidiary']}: 진행중 {row['Planned_Courses']}개, 완료 {row['Completed_Courses']}개, 완료율 {row['Course_Completion_Rate']:.2f}%, 계획시간 {row['Planned_Hours']:.1f}시간, 수강시간 {row['Actual_Hours']:.1f}시간, 이수율 {row['Hours_Completion_Rate']:.2f}%{region_info}")
                    else:
                        logger.info(f"  {row['Subsidiary']}: 진행중 {row['Planned_Courses']}개, 완료 {row['Completed_Courses']}개, 완료율 {row['Course_Completion_Rate']:.2f}%{region_info}")

                # 8.2단계: 전체 단계 결과 통합 (향후 확장용)
                logger.info("8.2단계: 전체 단계 결과를 통합합니다...")

                final_data = {
                    'step3_lms_completion': step3_result,
                    'step4_monthly_learning': step4_result,
                    'step5_new_hire': step5_result,
                    'step6_hipo': step6_result,
                    'step7_new_leader': step7_result,
                    'step8_logic_file': output_path
                }

                logger.info("✓ 최종 데이터 통합 완료")
                logger.info("✓ 포함된 단계:")
                logger.info("  - 3단계: LMS 과정등록율")
                logger.info("  - 4단계: 월별 학습시간")
                logger.info("  - 5단계: 신입사원 교육 이수율")
                logger.info("  - 6단계: 핵심인재 교육 이수율")
                logger.info("  - 7단계: 신입 팀장 교육 이수율")
                logger.info("  - 8단계: 최종 logic.csv 파일")

                return final_data
            else:
                logger.error("✗ 3단계 결과 데이터가 없습니다.")
                return None
        else:
            logger.error("✗ 3단계 결과가 없습니다.")
            return None

    except Exception as e:
        logger.error(f"✗ 오류 발생: {e}")
        return None

def run_make_logic(file_directory, analysis_year, analysis_month):
    """
    로직 생성 메인 실행 함수

    Args:
        file_directory (str): 파일이 있는 디렉토리
        analysis_year (int): 분석 기준 년도
        analysis_month (int): 분석 기준 월
    """
    # 전역 변수 설정 (모든 함수에서 사용)
    global ANALYSIS_YEAR, ANALYSIS_MONTH, ANALYSIS_MONTH_STR
    ANALYSIS_YEAR = analysis_year
    ANALYSIS_MONTH = analysis_month
    ANALYSIS_MONTH_STR = f"{ANALYSIS_YEAR}-{ANALYSIS_MONTH:02d}"

    logger.info("=== 로직 생성 시스템 시작 ===")
    logger.info(f"분석 기준: {ANALYSIS_YEAR}년 {ANALYSIS_MONTH}월")

    # 1단계: 전처리된 파일들 불러오기
    processed_files = load_processed_files(file_directory)

    if processed_files is not None:
        logger.info("✓ 로직 생성 준비 완료")

        # 2단계: 조인 테이블 생성
        logger.info("2단계: 조인 테이블 생성을 시작합니다...")

        # 2단계: HR과 LMS 테이블 조인
        join_table = create_join_table(processed_files['hr'], processed_files['lms'], file_directory)

        if join_table is not None:
            logger.info("✓ 2단계 완료")
        else:
            logger.error("✗ 2단계 실패")
            return False

        # 3단계: LMS 과정등록율 계산을 위한 과정
        logger.info("3단계: LMS 과정등록율 계산을 위한 과정을 시작합니다...")

        # 3.1단계: 이번달 교육계획 개수 계산
        current_education_result = calculate_current_education_plans(processed_files['hong_plan'])

        if current_education_result is not None:
            logger.info("✓ 3.1단계 완료")

            # 3.2단계: Final Sub.별 완료된 과정 개수 계산
            completed_courses_result = calculate_completed_courses_by_subsidiary(join_table)

            if completed_courses_result is not None:
                logger.info("✓ 3.2단계 완료")

                # 3.3단계: 완료율 계산
                completion_rate_result = calculate_completion_rate(current_education_result, completed_courses_result)

                if completion_rate_result is not None:
                    logger.info("✓ 3.3단계 완료")
                else:
                    logger.error("✗ 3.3단계 실패")
                    return False
        else:
            logger.error("✗ 3.1단계 실패")
            return False

        # 4단계: 월별 법인별 Learning Hrs. 계획 시간 계산
        monthly_learning_result = calculate_monthly_learning_hours(processed_files['hong_plan'])

        if monthly_learning_result is not None:
            logger.info("✓ 4.1단계 완료")

            # 4.2단계: 법인별 월별 실제 수강 시간 계산
            monthly_actual_result = calculate_monthly_actual_hours(join_table)

            if monthly_actual_result is not None:
                logger.info("✓ 4.2단계 완료")
            else:
                logger.error("✗ 4.2단계 실패")
                return False
        else:
            logger.error("✗ 4.1단계 실패")
            return False

        # 4.3단계: 법인별 월별 이수율 계산
        monthly_completion_result = calculate_monthly_completion_rate(monthly_learning_result, monthly_actual_result)

        if monthly_completion_result is not None:
            logger.info("✓ 4.3단계 완료")
        else:
            logger.error("✗ 4.3단계 실패")
            return False

        # 5단계: 신입사원 교육 이수율 계산
        new_hire_result = calculate_new_hire_completion_rate(join_table)

        if new_hire_result is not None:
            logger.info("✓ 5단계 완료")
        else:
            logger.error("✗ 5단계 실패")
            return False

        # 6단계: 핵심인재 교육 이수율 계산
        hipo_result = calculate_hipo_completion_rate(join_table)

        if hipo_result is not None:
            logger.info("✓ 6단계 완료")
        else:
            logger.error("✗ 6단계 실패")
            return False

        # 7단계: 신입 팀장 교육 이수율 계산
        new_leader_result = calculate_new_leader_completion_rate(join_table)

        if new_leader_result is not None:
            logger.info("✓ 7단계 완료")
        else:
            logger.error("✗ 7단계 실패")
            return False

        # 8단계: 최종 데이터 생성
        final_result = create_final_logic_data(completion_rate_result, monthly_completion_result, new_hire_result, hipo_result, new_leader_result, file_directory)

        if final_result is not None:
            logger.info("✓ 8단계 완료")
        else:
            logger.error("✗ 8단계 실패")
            return False

        logger.info("=== 로직 생성 시스템 완료 ===")
        return True
    else:
        logger.error("✗ 전처리된 파일 불러오기에 실패했습니다.")
        return False

if __name__ == "__main__":
    # 직접 실행 시 기본값 사용
    default_year = 2025
    default_month = 8
    default_directory = f"data/{default_month}"

    print(f"직접 실행 모드: {default_year}년 {default_month}월 데이터 분석")
    success = run_make_logic(default_directory, default_year, default_month)
    sys.exit(0 if success else 1)
