# -*- encoding: utf-8 -*-
"""
Data caching module for subsidiary and other data
"""

import os
import pandas as pd
from pathlib import Path

class DataCache:
    """데이터 캐싱 클래스"""

    def __init__(self):
        self.subsidiary_cache = {}
        self.region_cache = {}
        self.logic_region_cache = {}
        self.logic_global_cache = {}
        # 상대 경로로 변경 (실행 위치 기준)
        # __file__은 apps/data_cache.py이므로 parent.parent가 프로젝트 root
        base_dir = Path(__file__).resolve().parent.parent
        self.base_path = base_dir / "data"
        self._load_subsidiary_data()
        self._load_region_data()
        self._load_logic_region_data()
        self._load_logic_global_data()

    def _load_subsidiary_data(self):
        """모든 월의 subsidiary 데이터를 로드"""
        print("Loading subsidiary data for all months...")

        for month in range(1, 13):
            month_folder = str(month)
            csv_file = self.base_path / month_folder / "hr_index_final.csv"

            if csv_file.exists():
                try:
                    df = pd.read_csv(csv_file)
                    if 'Final Sub.' in df.columns:
                        # NaN 값 제거하고 unique 값만 추출
                        subsidiaries = df['Final Sub.'].dropna().unique().tolist()
                        # 빈 문자열 제거
                        subsidiaries = [s for s in subsidiaries if s and str(s).strip()]
                        self.subsidiary_cache[month] = sorted(subsidiaries)
                        print(f"Loaded {len(subsidiaries)} subsidiaries for {month}월")
                    else:
                        print(f"Warning: 'Final Sub.' column not found in {csv_file}")
                        self.subsidiary_cache[month] = []
                except Exception as e:
                    print(f"Error loading {csv_file}: {e}")
                    self.subsidiary_cache[month] = []
            else:
                print(f"File not found: {csv_file}")
                self.subsidiary_cache[month] = []

        print(f"Subsidiary cache loaded for {len(self.subsidiary_cache)} months")

    def _load_region_data(self):
        """모든 월의 region 데이터를 로드"""
        print("Loading region data for all months...")

        for month in range(1, 13):
            month_folder = str(month)
            csv_file = self.base_path / month_folder / "hr_index_final.csv"

            if csv_file.exists():
                try:
                    df = pd.read_csv(csv_file)
                    if 'Final Region' in df.columns:
                        # NaN 값 제거하고 unique 값만 추출
                        regions = df['Final Region'].dropna().unique().tolist()
                        # 빈 문자열 제거
                        regions = [r for r in regions if r and str(r).strip()]
                        self.region_cache[month] = sorted(regions)
                        print(f"Loaded {len(regions)} regions for {month}월")
                    else:
                        print(f"Warning: 'Final Region' column not found in {csv_file}")
                        self.region_cache[month] = []
                except Exception as e:
                    print(f"Error loading {csv_file}: {e}")
                    self.region_cache[month] = []
            else:
                print(f"File not found: {csv_file}")
                self.region_cache[month] = []

        print(f"Region cache loaded for {len(self.region_cache)} months")

    def _load_logic_region_data(self):
        """모든 월의 logic.csv 지역별 데이터를 로드"""
        print("Loading logic region data for all months...")

        for month in range(1, 13):
            month_folder = str(month)
            csv_file = self.base_path / month_folder / "logic.csv"

            if csv_file.exists():
                try:
                    df = pd.read_csv(csv_file)
                    print(f"Debug - {month}월 logic.csv 로드 완료: {len(df)} 행")

                    # Final Region 컬럼이 있는지 확인
                    if 'Final Region' not in df.columns:
                        print(f"Warning: 'Final Region' column not found in {csv_file}")
                        self.logic_region_cache[month] = {}
                        continue

                    # sum할 컬럼들 정의
                    sum_columns = [
                        'Planned_Courses', 'Completed_Courses', 'Planned_Hours', 'Actual_Hours',
                        'New_Hire_Completed', 'New_Hire_Not_Completed', 'New_Hire_Pending', 'New_Hire_Total',
                        'EIP_Completed', 'EIP_Not_Completed', 'EIP_Total',
                        'GLP_Completed', 'GLP_Not_Completed', 'GLP_Total',
                        'New_Leader_Completed', 'New_Leader_Not_Completed', 'New_Leader_Total'
                    ]

                    # 존재하는 컬럼만 선택
                    available_columns = [col for col in sum_columns if col in df.columns]
                    print(f"Debug - {month}월 사용 가능한 컬럼: {available_columns}")

                    # Final Region으로 그룹핑하고 sum 계산
                    region_groups = df.groupby('Final Region')[available_columns].sum()

                    # 모든 컬럼 정의 (29개 컬럼: 기존 25개 + New_Leader 4개)
                    all_columns = [
                        'Planned_Courses', 'Completed_Courses', 'Planned_Hours', 'Actual_Hours',
                        'New_Hire_Completed', 'New_Hire_Not_Completed', 'New_Hire_Pending', 'New_Hire_Total',
                        'EIP_Completed', 'EIP_Not_Completed', 'EIP_Total',
                        'GLP_Completed', 'GLP_Not_Completed', 'GLP_Total',
                        'New_Leader_Completed', 'New_Leader_Not_Completed', 'New_Leader_Total',
                        'Course_Completion_Rate', 'Hours_Completion_Rate', 'New_Hire_Completion_Rate',
                        'EIP_Completion_Rate', 'GLP_Completion_Rate', 'New_Leader_Completion_Rate',
                        'JAM_Member_Rate', 'Annual_Plan_Setup_Rate', 'New_LMS_Course_Rate',
                        'LMS_Mission_Rate', 'Global_LD_Council_Rate', 'Infra_Index_Response_Rate'
                    ]

                    # 각 지역별로 데이터 처리
                    month_data = {}
                    for region in region_groups.index:
                        # 기본 sum 데이터 가져오기
                        region_sum_data = region_groups.loc[region].to_dict()

                        # 모든 컬럼을 0으로 초기화
                        region_data = {col: 0 for col in all_columns}

                        # sum 컬럼들 값을 설정
                        for col in available_columns:
                            if col in region_sum_data:
                                region_data[col] = region_sum_data[col]

                        # 비율 계산 (소수 둘째자리에서 반올림)
                        # Course_Completion_Rate = Completed_Courses / Planned_Courses * 100
                        if region_data['Planned_Courses'] > 0:
                            region_data['Course_Completion_Rate'] = round(region_data['Completed_Courses'] / region_data['Planned_Courses'] * 100, 2)

                        # Hours_Completion_Rate = Actual_Hours / Planned_Hours * 100
                        if region_data['Planned_Hours'] > 0:
                            region_data['Hours_Completion_Rate'] = round(region_data['Actual_Hours'] / region_data['Planned_Hours'] * 100, 2)

                        # New_Hire_Completion_Rate = (New_Hire_Completed + New_Hire_Pending) / New_Hire_Total * 100
                        if region_data['New_Hire_Total'] > 0:
                            region_data['New_Hire_Completion_Rate'] = round((region_data['New_Hire_Completed'] + region_data['New_Hire_Pending']) / region_data['New_Hire_Total'] * 100, 2)

                        # EIP_Completion_Rate = EIP_Completed / EIP_Total * 100
                        if region_data['EIP_Total'] > 0:
                            region_data['EIP_Completion_Rate'] = round(region_data['EIP_Completed'] / region_data['EIP_Total'] * 100, 2)

                        # GLP_Completion_Rate = GLP_Completed / GLP_Total * 100
                        if region_data['GLP_Total'] > 0:
                            region_data['GLP_Completion_Rate'] = round(region_data['GLP_Completed'] / region_data['GLP_Total'] * 100, 2)

                        # New_Leader_Completion_Rate = New_Leader_Completed / New_Leader_Total * 100
                        if region_data['New_Leader_Total'] > 0:
                            region_data['New_Leader_Completion_Rate'] = round(region_data['New_Leader_Completed'] / region_data['New_Leader_Total'] * 100, 2)

                        # 6개 Y/N 컬럼의 Y 비율 계산
                        # 해당 지역의 법인 수
                        region_df = df[df['Final Region'] == region]
                        total_subsidiaries = len(region_df)

                        if total_subsidiaries > 0:
                            # JAM Member
                            if 'JAM Member' in df.columns:
                                y_count = len(region_df[region_df['JAM Member'].astype(str).str.upper() == 'Y'])
                                region_data['JAM_Member_Rate'] = round(y_count / total_subsidiaries * 100, 2)

                            # Annual Plan Setup
                            if 'Annual Plan Setup' in df.columns:
                                y_count = len(region_df[region_df['Annual Plan Setup'].astype(str).str.upper() == 'Y'])
                                region_data['Annual_Plan_Setup_Rate'] = round(y_count / total_subsidiaries * 100, 2)

                            # New LMS Course
                            if 'New LMS Course' in df.columns:
                                y_count = len(region_df[region_df['New LMS Course'].astype(str).str.upper() == 'Y'])
                                region_data['New_LMS_Course_Rate'] = round(y_count / total_subsidiaries * 100, 2)

                            # LMS Mission
                            if 'LMS Mission' in df.columns:
                                y_count = len(region_df[region_df['LMS Mission'].astype(str).str.upper() == 'Y'])
                                region_data['LMS_Mission_Rate'] = round(y_count / total_subsidiaries * 100, 2)

                            # Global L&D Council
                            if 'Global L&D Council' in df.columns:
                                y_count = len(region_df[region_df['Global L&D Council'].astype(str).str.upper() == 'Y'])
                                region_data['Global_LD_Council_Rate'] = round(y_count / total_subsidiaries * 100, 2)

                            # Infra index response
                            if 'Infra index response' in df.columns:
                                y_count = len(region_df[region_df['Infra index response'].astype(str).str.upper() == 'Y'])
                                region_data['Infra_Index_Response_Rate'] = round(y_count / total_subsidiaries * 100, 2)

                        # Score 평균 계산
                        if 'Score' in df.columns:
                            region_scores = region_df['Score'].dropna()
                            if len(region_scores) > 0:
                                region_data['Score'] = round(region_scores.mean(), 2)
                            else:
                                region_data['Score'] = 0
                        else:
                            region_data['Score'] = 0

                        month_data[region] = region_data
                        print(f"Debug - {month}월 {region} 지역 데이터: {len(region_data)} 컬럼 (모든 컬럼 포함, Score 평균: {region_data.get('Score', 0):.2f}점)")

                    self.logic_region_cache[month] = month_data
                    print(f"Debug - {month}월 logic 지역 캐시 완료: {len(month_data)} 지역")

                except Exception as e:
                    print(f"Error loading {csv_file}: {e}")
                    self.logic_region_cache[month] = {}
            else:
                print(f"File not found: {csv_file}")
                self.logic_region_cache[month] = {}

        print(f"Logic region cache loaded for {len(self.logic_region_cache)} months")

    def _load_logic_global_data(self):
        """모든 월의 logic 데이터를 Global 단위로 로드"""
        print("Loading logic global data for all months...")

        for month in range(1, 13):
            month_folder = str(month)
            csv_file = self.base_path / month_folder / "logic.csv"

            if csv_file.exists():
                try:
                    df = pd.read_csv(csv_file)
                    print(f"Debug - {month}월 logic.csv 로드 완료: {len(df)} 행")

                    # 모든 데이터를 하나의 그룹으로 합치기 (Global)
                    # 합계할 컬럼들
                    sum_columns = [
                        'Planned_Courses', 'Completed_Courses', 'Planned_Hours', 'Actual_Hours',
                        'New_Hire_Completed', 'New_Hire_Not_Completed', 'New_Hire_Pending', 'New_Hire_Total',
                        'EIP_Completed', 'EIP_Not_Completed', 'EIP_Total',
                        'GLP_Completed', 'GLP_Not_Completed', 'GLP_Total',
                        'New_Leader_Completed', 'New_Leader_Not_Completed', 'New_Leader_Total'
                    ]

                    # Global 데이터 초기화 (모든 컬럼을 0으로 설정)
                    global_data = {}
                    for col in sum_columns:
                        global_data[col] = 0

                    # 실제 데이터가 있는 경우에만 합계 계산
                    if not df.empty:
                        global_sum_data = df[sum_columns].sum()
                        for col in sum_columns:
                            global_data[col] = global_sum_data[col]

                    # 비율 계산 (소수 둘째자리에서 반올림)
                    # Course_Completion_Rate = Completed_Courses / Planned_Courses * 100
                    if global_data['Planned_Courses'] > 0:
                        global_data['Course_Completion_Rate'] = round(global_data['Completed_Courses'] / global_data['Planned_Courses'] * 100, 2)
                    else:
                        global_data['Course_Completion_Rate'] = 0

                    # Hours_Completion_Rate = Actual_Hours / Planned_Hours * 100
                    if global_data['Planned_Hours'] > 0:
                        global_data['Hours_Completion_Rate'] = round(global_data['Actual_Hours'] / global_data['Planned_Hours'] * 100, 2)
                    else:
                        global_data['Hours_Completion_Rate'] = 0

                    # New_Hire_Completion_Rate = (New_Hire_Completed + New_Hire_Pending) / New_Hire_Total * 100
                    if global_data['New_Hire_Total'] > 0:
                        global_data['New_Hire_Completion_Rate'] = round((global_data['New_Hire_Completed'] + global_data['New_Hire_Pending']) / global_data['New_Hire_Total'] * 100, 2)
                    else:
                        global_data['New_Hire_Completion_Rate'] = 0

                    # EIP_Completion_Rate = EIP_Completed / EIP_Total * 100
                    if global_data['EIP_Total'] > 0:
                        global_data['EIP_Completion_Rate'] = round(global_data['EIP_Completed'] / global_data['EIP_Total'] * 100, 2)
                    else:
                        global_data['EIP_Completion_Rate'] = 0

                    # GLP_Completion_Rate = GLP_Completed / GLP_Total * 100
                    if global_data['GLP_Total'] > 0:
                        global_data['GLP_Completion_Rate'] = round(global_data['GLP_Completed'] / global_data['GLP_Total'] * 100, 2)
                    else:
                        global_data['GLP_Completion_Rate'] = 0

                    # New_Leader_Completion_Rate = New_Leader_Completed / New_Leader_Total * 100
                    if global_data['New_Leader_Total'] > 0:
                        global_data['New_Leader_Completion_Rate'] = round(global_data['New_Leader_Completed'] / global_data['New_Leader_Total'] * 100, 2)
                    else:
                        global_data['New_Leader_Completion_Rate'] = 0

                    # 6개 Y/N 컬럼의 Y 비율 계산 (전체 법인 기준)
                    total_subsidiaries = len(df)

                    if total_subsidiaries > 0:
                        # JAM Member
                        if 'JAM Member' in df.columns:
                            y_count = len(df[df['JAM Member'].astype(str).str.upper() == 'Y'])
                            global_data['JAM_Member_Rate'] = round(y_count / total_subsidiaries * 100, 2)
                        else:
                            global_data['JAM_Member_Rate'] = 0

                        # Annual Plan Setup
                        if 'Annual Plan Setup' in df.columns:
                            y_count = len(df[df['Annual Plan Setup'].astype(str).str.upper() == 'Y'])
                            global_data['Annual_Plan_Setup_Rate'] = round(y_count / total_subsidiaries * 100, 2)
                        else:
                            global_data['Annual_Plan_Setup_Rate'] = 0

                        # New LMS Course
                        if 'New LMS Course' in df.columns:
                            y_count = len(df[df['New LMS Course'].astype(str).str.upper() == 'Y'])
                            global_data['New_LMS_Course_Rate'] = round(y_count / total_subsidiaries * 100, 2)
                        else:
                            global_data['New_LMS_Course_Rate'] = 0

                        # LMS Mission
                        if 'LMS Mission' in df.columns:
                            y_count = len(df[df['LMS Mission'].astype(str).str.upper() == 'Y'])
                            global_data['LMS_Mission_Rate'] = round(y_count / total_subsidiaries * 100, 2)
                        else:
                            global_data['LMS_Mission_Rate'] = 0

                        # Global L&D Council
                        if 'Global L&D Council' in df.columns:
                            y_count = len(df[df['Global L&D Council'].astype(str).str.upper() == 'Y'])
                            global_data['Global_LD_Council_Rate'] = round(y_count / total_subsidiaries * 100, 2)
                        else:
                            global_data['Global_LD_Council_Rate'] = 0

                        # Infra index response
                        if 'Infra index response' in df.columns:
                            y_count = len(df[df['Infra index response'].astype(str).str.upper() == 'Y'])
                            global_data['Infra_Index_Response_Rate'] = round(y_count / total_subsidiaries * 100, 2)
                        else:
                            global_data['Infra_Index_Response_Rate'] = 0

                    # Score 평균 계산
                    if 'Score' in df.columns:
                        global_scores = df['Score'].dropna()
                        if len(global_scores) > 0:
                            global_data['Score'] = round(global_scores.mean(), 2)
                        else:
                            global_data['Score'] = 0
                    else:
                        global_data['Score'] = 0

                    self.logic_global_cache[month] = global_data
                    print(f"Debug - {month}월 logic global 캐시 완료: {len(global_data)} 컬럼 (Score 평균: {global_data.get('Score', 0):.2f}점)")

                except Exception as e:
                    print(f"Error loading {csv_file}: {e}")
                    self.logic_global_cache[month] = {}
            else:
                print(f"File not found: {csv_file}")
                self.logic_global_cache[month] = {}

        print(f"Logic global cache loaded for {len(self.logic_global_cache)} months")

    def get_subsidiaries_by_month(self, month):
        """특정 월의 subsidiary 목록 반환"""
        # 캐시에 없으면 파일에서 직접 읽기
        if month not in self.subsidiary_cache:
            self._load_month_data(month)
        return self.subsidiary_cache.get(month, [])

    def get_regions_by_month(self, month):
        """특정 월의 region 목록 반환"""
        # 캐시에 없으면 파일에서 직접 읽기
        if month not in self.region_cache:
            self._load_month_data(month)
        return self.region_cache.get(month, [])

    def get_all_months_with_data(self):
        """데이터가 있는 모든 월 목록 반환"""
        return [month for month, data in self.subsidiary_cache.items() if data]

    def has_data_for_month(self, month):
        """특정 월에 데이터가 있는지 확인"""
        return month in self.subsidiary_cache and len(self.subsidiary_cache[month]) > 0

    def get_logic_region_data(self, month, region):
        """특정 월과 지역의 logic 데이터 반환"""
        if month in self.logic_region_cache and region in self.logic_region_cache[month]:
            return self.logic_region_cache[month][region]
        return None

    def get_logic_global_data(self, month):
        """특정 월의 Global logic 데이터 반환"""
        if month in self.logic_global_cache:
            return self.logic_global_cache[month]
        return None

    def get_subsidiary_summary_data(self, month):
        """특정 월의 Subsidiary 요약 데이터 반환"""
        month_folder = str(month)
        csv_file = self.base_path / month_folder / "hr_index_final.csv"

        if not csv_file.exists():
            return []

        try:
            df = pd.read_csv(csv_file)

            # 필요한 컬럼들이 있는지 확인
            required_columns = ['Final Sub.', 'Final Region', 'New Hire', 'HIPO Type', 'Staff/Operator']
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                print(f"Warning: Missing columns in {csv_file}: {missing_columns}")
                return []

            # Final Sub.로 그룹핑
            grouped = df.groupby('Final Sub.').agg({
                'Final Region': 'max',  # max(Final Region)
                'New Hire': 'count',   # 전체 count
                'HIPO Type': lambda x: (x == 'EIP').sum(),  # EIP count
                'Staff/Operator': lambda x: (x == 'Staff').sum()  # Staff count
            }).reset_index()

            # GLP count 계산
            glp_count = df[df['HIPO Type'] == 'GLP'].groupby('Final Sub.').size().reset_index(name='GLP Count')
            grouped = grouped.merge(glp_count, left_on='Final Sub.', right_on='Final Sub.', how='left')
            grouped['GLP Count'] = grouped['GLP Count'].fillna(0).astype(int)

            # Operator count 계산
            grouped['Operator Count'] = grouped['New Hire'] - grouped['Staff/Operator']

            # 컬럼명 정리
            grouped.columns = ['Subsidiary', 'Region', 'Total Count', 'EIP Count', 'Staff Count', 'GLP Count', 'Operator Count']

            # New Hire가 Y인 것만 count
            new_hire_count = df[df['New Hire'] == 'Y'].groupby('Final Sub.').size().reset_index(name='New Hire Count')
            grouped = grouped.merge(new_hire_count, left_on='Subsidiary', right_on='Final Sub.', how='left')
            grouped = grouped.drop('Final Sub.', axis=1)
            grouped['New Hire Count'] = grouped['New Hire Count'].fillna(0).astype(int)

            # 최종 컬럼 순서 정리
            result = grouped[['Subsidiary', 'Region', 'Total Count', 'New Hire Count', 'EIP Count', 'GLP Count', 'Staff Count', 'Operator Count']]

            return result.to_dict('records')

        except Exception as e:
            print(f"Error processing {csv_file}: {e}")
            return []

    def _load_month_data(self, month):
        """특정 월의 데이터만 로드 (캐시에 없을 때)"""
        month_folder = str(month)
        csv_file = self.base_path / month_folder / "hr_index_final.csv"

        if csv_file.exists():
            try:
                df = pd.read_csv(csv_file)

                # Subsidiary 로드
                if 'Final Sub.' in df.columns:
                    subsidiaries = df['Final Sub.'].dropna().unique().tolist()
                    subsidiaries = [s for s in subsidiaries if s and str(s).strip()]
                    self.subsidiary_cache[month] = sorted(subsidiaries)

                # Region 로드
                if 'Final Region' in df.columns:
                    regions = df['Final Region'].dropna().unique().tolist()
                    regions = [r for r in regions if r and str(r).strip()]
                    self.region_cache[month] = sorted(regions)

                print(f"Loaded data for {month}월 from file")
            except Exception as e:
                print(f"Error loading data for {month}월: {e}")

    def reload_cache(self):
        """캐시 재로드"""
        self.subsidiary_cache = {}
        self.region_cache = {}
        self.logic_region_cache = {}
        self.logic_global_cache = {}
        self._load_subsidiary_data()
        self._load_region_data()
        self._load_logic_region_data()
        self._load_logic_global_data()

# 전역 인스턴스 생성
data_cache = DataCache()
