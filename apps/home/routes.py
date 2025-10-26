# -*- encoding: utf-8 -*-
"""
Copyright (c) 2019 - present AppSeed.us
"""

from apps.home import blueprint
from flask import render_template, request, jsonify
from flask_login import login_required
from jinja2 import TemplateNotFound
from apps.data_cache import data_cache
from pathlib import Path

# 상대 경로 설정 (실행 위치 기준)
# __file__은 apps/home/routes.py이므로 parent.parent는 apps, 한 번 더 parent가 프로젝트 root
BASE_DIR = Path(__file__).resolve().parent.parent.parent
DATA_DIR = BASE_DIR / "data"


@blueprint.route('/index')
@login_required
def index():

    return render_template('home/index.html', segment='index')


@blueprint.route('/global')
@login_required
def global_page():
    """Global 페이지"""
    # 사용자가 선택한 기준월 (쿼리 파라미터에서 가져오기)
    current_month = request.args.get('month', 9)  # 기본값은 9월

    # Global 지표점수 데이터 가져오기
    global_logic_data = data_cache.get_logic_global_data(int(current_month))

    if global_logic_data:
        print(f"\n=== {current_month}월 Global 지표점수 ===")
        print(f"LMS 과정 등록률: {global_logic_data.get('Course_Completion_Rate', 0):.2f}%")
        print(f"계획대비 실행률: {global_logic_data.get('Hours_Completion_Rate', 0):.2f}%")
        print(f"신규입사자 교육 이수율: {global_logic_data.get('New_Hire_Completion_Rate', 0):.2f}%")
        print(f"핵심인재(EIP) 교육 이수율: {global_logic_data.get('EIP_Completion_Rate', 0):.2f}%")
        print(f"핵심인재(GLP) 교육 이수율: {global_logic_data.get('GLP_Completion_Rate', 0):.2f}%")
        print(f"=== 상세 데이터 ===")
        print(f"계획된 과정: {global_logic_data.get('Planned_Courses', 0):,}개")
        print(f"완료된 과정: {global_logic_data.get('Completed_Courses', 0):,}개")
        print(f"계획된 시간: {global_logic_data.get('Planned_Hours', 0):,}시간")
        print(f"실제 시간: {global_logic_data.get('Actual_Hours', 0):,}시간")
        print(f"신규입사자 총원: {global_logic_data.get('New_Hire_Total', 0):,}명")
        print(f"신규입사자 완료: {global_logic_data.get('New_Hire_Completed', 0):,}명")
        print(f"EIP 총원: {global_logic_data.get('EIP_Total', 0):,}명")
        print(f"EIP 완료: {global_logic_data.get('EIP_Completed', 0):,}명")
        print(f"GLP 총원: {global_logic_data.get('GLP_Total', 0):,}명")
        print(f"GLP 완료: {global_logic_data.get('GLP_Completed', 0):,}명")
        print("=" * 50)
    else:
        print(f"\n=== {current_month}월 Global 지표점수 ===")
        print("Global 데이터를 찾을 수 없습니다.")
        print("=" * 50)

    return render_template('home/global.html', segment='global', current_month=int(current_month))


@blueprint.route('/region')
@login_required
def region_page():
    """Region 페이지"""
    # 사용자가 선택한 기준월과 선택된 region (쿼리 파라미터에서 가져오기)
    current_month = request.args.get('month', 9)  # 기본값은 9월
    selected_region = request.args.get('selected')  # 선택된 region

    # 데이터 존재 여부 확인
    has_data = data_cache.has_data_for_month(int(current_month))

    return render_template('home/region.html',
                         segment='region',
                         current_month=current_month,
                         selected_region=selected_region,
                         has_data=has_data)


@blueprint.route('/region/<region>')
@login_required
def region_detail_page(region):
    """Region Detail 페이지"""
    current_month = request.args.get('month', type=int)

    if current_month is None:
        abort(404)

    has_data = data_cache.has_data_for_month(current_month)

    return render_template('home/region_detail.html',
                         region=region,
                         current_month=current_month,
                         has_data=has_data)


@blueprint.route('/subsidiary')
@login_required
def subsidiary_page():
    """Subsidiary 페이지"""
    # 사용자가 선택한 기준월 (쿼리 파라미터에서 가져오기)
    current_month = request.args.get('month', 9)  # 기본값은 9월

    return render_template('home/subsidiary.html',
                         segment='subsidiary',
                         current_month=current_month)


@blueprint.route('/subsidiary/<subsidiary>')
@login_required
def subsidiary_detail_page(subsidiary):
    """특정 Subsidiary 상세 페이지"""
    # 사용자가 선택한 기준월 (쿼리 파라미터에서 가져오기)
    current_month = request.args.get('month', 9)  # 기본값 9월

    # 데이터 존재 여부 확인
    has_data = data_cache.has_data_for_month(int(current_month))

    return render_template('home/subsidiary_detail.html',
                         segment='subsidiary_detail',
                         subsidiary=subsidiary,
                         current_month=current_month,
                         has_data=has_data)


@blueprint.route('/api/subsidiaries/<int:month>')
@login_required
def get_subsidiaries(month):
    """특정 월의 subsidiary 목록을 반환하는 API"""
    try:
        subsidiaries = data_cache.get_subsidiaries_by_month(month)
        return jsonify({
            'success': True,
            'month': month,
            'subsidiaries': subsidiaries
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@blueprint.route('/api/regions/<int:month>')
@login_required
def get_regions(month):
    """특정 월의 region 목록을 반환하는 API"""
    try:
        regions = data_cache.get_regions_by_month(month)
        return jsonify({
            'success': True,
            'month': month,
            'regions': regions
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@blueprint.route('/api/subsidiary-summary/<int:month>')
@login_required
def get_subsidiary_summary(month):
    """특정 월의 Subsidiary 요약 데이터를 반환하는 API"""
    try:
        summary_data = data_cache.get_subsidiary_summary_data(month)
        return jsonify({
            'success': True,
            'month': month,
            'data': summary_data
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@blueprint.route('/api/subsidiary-detail/<subsidiary>/<int:month>')
@login_required
def get_subsidiary_detail(subsidiary, month):
    """특정 법인의 상세 정보를 반환하는 API"""
    try:
        import pandas as pd
        from pathlib import Path

        # hr_index_final.csv 파일 경로
        month_folder = str(month)
        csv_file = DATA_DIR / month_folder / "hr_index_final.csv"

        if not csv_file.exists():
            return jsonify({
                'success': False,
                'error': f'{month_folder} 데이터가 존재하지 않습니다.'
            }), 404

        # CSV 파일 읽기
        df = pd.read_csv(csv_file)

        # Final Sub. 컬럼이 있는지 확인
        if 'Final Sub.' not in df.columns:
            return jsonify({
                'success': False,
                'error': 'Final Sub. 컬럼을 찾을 수 없습니다.'
            }), 400

        # 해당 subsidiary 데이터 필터링
        subsidiary_data = df[df['Final Sub.'] == subsidiary]

        if subsidiary_data.empty:
            return jsonify({
                'success': False,
                'error': f'{subsidiary} 법인의 데이터를 찾을 수 없습니다.'
            }), 404

        # 기본 정보 계산
        total_count = len(subsidiary_data)
        region = subsidiary_data['Final Region'].iloc[0] if 'Final Region' in subsidiary_data.columns else '-'

        # 법인명과 Sub. Name(MP) 가져오기
        subsidiary_name = subsidiary
        sub_name_mp = subsidiary_data['Sub. Name(MP)'].iloc[0] if 'Sub. Name(MP)' in subsidiary_data.columns else ''

        # 디버깅 정보 출력 (필요시 주석 해제)
        # print(f"Debug - subsidiary: {subsidiary}")
        # print(f"Debug - subsidiary_data shape: {subsidiary_data.shape}")
        # print(f"Debug - columns: {list(subsidiary_data.columns)}")
        # print(f"Debug - region: {region}")
        # print(f"Debug - sub_name_mp: {sub_name_mp}")

        # New Hire Count (Y인 것만)
        new_hire_count = len(subsidiary_data[subsidiary_data['New Hire'] == 'Y']) if 'New Hire' in subsidiary_data.columns else 0

        # EIP Count
        eip_count = len(subsidiary_data[subsidiary_data['HIPO Type'] == 'EIP']) if 'HIPO Type' in subsidiary_data.columns else 0

        # GLP Count
        glp_count = len(subsidiary_data[subsidiary_data['HIPO Type'] == 'GLP']) if 'HIPO Type' in subsidiary_data.columns else 0

        # Staff Count
        staff_count = len(subsidiary_data[subsidiary_data['Staff/Operator'] == 'Staff']) if 'Staff/Operator' in subsidiary_data.columns else 0

        # Operator Count
        operator_count = len(subsidiary_data[subsidiary_data['Staff/Operator'] == 'Operator']) if 'Staff/Operator' in subsidiary_data.columns else 0

        # 담당자 정보 가져오기 (hong_data_manager_final.csv)
        manager_email = '-'
        try:
            manager_file = DATA_DIR / month_folder / "hong_data_manager_final.csv"

            if manager_file.exists():
                manager_df = pd.read_csv(manager_file)

                if 'Final Sub.' in manager_df.columns and 'L&D PIC e-mail' in manager_df.columns:
                    manager_data = manager_df[manager_df['Final Sub.'] == subsidiary]

                    if not manager_data.empty:
                        manager_email = manager_data['L&D PIC e-mail'].iloc[0]
                        if pd.isna(manager_email):
                            manager_email = '-'
        except Exception as e:
            print(f"Error loading manager data: {e}")
            manager_email = '-'

        response_data = {
            'success': True,
            'subsidiary': subsidiary,
            'month': month,
            'data': {
                'region': region,
                'total_count': total_count,
                'new_hire_count': new_hire_count,
                'eip_count': eip_count,
                'glp_count': glp_count,
                'staff_count': staff_count,
                'operator_count': operator_count
            },
            'company_info': {
                'subsidiary_name': subsidiary_name,
                'sub_name_mp': sub_name_mp,
                'region': region,
                'manager_email': manager_email
            }
        }

        return jsonify(response_data)

    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@blueprint.route('/api/region-logic-data/<region>/<int:month>')
@login_required
def get_region_logic_data(region, month):
    """특정 지역의 logic 데이터를 반환"""
    try:
        from apps.data_cache import data_cache
        from urllib.parse import unquote

        # URL 디코딩
        region = unquote(region)
        print(f"Debug - Region logic data request: {region}, month: {month}")

        # 캐시에서 지역 데이터 가져오기
        region_data = data_cache.get_logic_region_data(month, region)

        if region_data is None:
            return jsonify({
                'success': False,
                'error': f'No data found for region {region} in month {month}'
            }), 404

        print(f"Debug - Region logic data found: {len(region_data)} fields")
        print(f"Debug - Region logic data: {region_data}")

        # Global 데이터 가져오기
        global_data = data_cache.get_logic_global_data(month)

        # NaN 값 처리 함수
        def safe_convert(obj):
            """numpy 타입과 NaN을 JSON 직렬화 가능한 형태로 변환"""
            import numpy as np
            if isinstance(obj, (np.integer, np.int64)):
                return int(obj)
            elif isinstance(obj, (np.floating, np.float64)):
                if np.isnan(obj):
                    return 0  # NaN을 0으로 변환
                return float(obj)
            elif isinstance(obj, dict):
                return {key: safe_convert(value) for key, value in obj.items()}
            elif isinstance(obj, list):
                return [safe_convert(item) for item in obj]
            else:
                return obj

        response_data = {
            'success': True,
            'region': region,
            'month': month,
            'data': safe_convert(region_data),
            'global_data': safe_convert(global_data)
        }

        return jsonify(response_data)

    except Exception as e:
        print(f"Error getting region logic data: {e}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@blueprint.route('/api/global-logic-data/<int:month>')
@login_required
def get_global_logic_data(month):
    """Global logic 데이터를 반환"""
    try:
        from apps.data_cache import data_cache

        print(f"Debug - Global logic data request: month: {month}")

        # 캐시에서 Global 데이터 가져오기
        global_data = data_cache.get_logic_global_data(month)

        if global_data is None:
            return jsonify({
                'success': False,
                'error': f'No data found for month {month}'
            }), 404

        print(f"Debug - Global logic data found: {len(global_data)} fields")
        print(f"Debug - Global logic data: {global_data}")

        response_data = {
            'success': True,
            'month': month,
            'data': global_data
        }

        return jsonify(response_data)

    except Exception as e:
        print(f"Error getting global logic data: {e}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@blueprint.route('/api/global-region-metrics/<int:month>')
@login_required
def get_global_region_metrics(month):
    """모든 지역의 지표 데이터를 반환 (캐시 사용)"""
    try:
        from apps.data_cache import data_cache

        print(f"Debug - Global region metrics request: month: {month}")

        # 캐시에서 해당 월의 모든 region 데이터 가져오기
        if month not in data_cache.logic_region_cache:
            return jsonify({
                'success': False,
                'error': f'{month}월 데이터가 존재하지 않습니다.'
            }), 404

        region_cache = data_cache.logic_region_cache[month]
        print(f"Debug - {month}월 region 캐시: {len(region_cache)} 개 지역")

        # 각 region별 데이터를 리스트로 변환
        regions = []
        for region_name, region_data in region_cache.items():
            new_leader_rate = round(region_data.get('New_Leader_Completion_Rate', 0), 1)
            print(f"Debug - {region_name} New_Leader_Completion_Rate: {new_leader_rate}% (원본: {region_data.get('New_Leader_Completion_Rate')})")

            region_metrics = {
                'region': region_name,
                'total_score': round(region_data.get('Score', 0), 1),  # 캐시된 Region Score 평균
                'course_completion_rate': round(region_data.get('Course_Completion_Rate', 0), 1),
                'hours_completion_rate': round(region_data.get('Hours_Completion_Rate', 0), 1),
                'new_hire_completion_rate': round(region_data.get('New_Hire_Completion_Rate', 0), 1),
                'eip_completion_rate': round(region_data.get('EIP_Completion_Rate', 0), 1),
                'glp_completion_rate': round(region_data.get('GLP_Completion_Rate', 0), 1),
                'new_leader_rate': new_leader_rate,
                'lms_education_rate': round(region_data.get('New_LMS_Course_Rate', 0), 1),
                'lms_mission_rate': round(region_data.get('LMS_Mission_Rate', 0), 1),
                'annual_plan': round(region_data.get('Annual_Plan_Setup_Rate', 0), 1),
                'jam_community': round(region_data.get('JAM_Member_Rate', 0), 1),
                'global_council': round(region_data.get('Global_LD_Council_Rate', 0), 1),
                'infra_index': round(region_data.get('Infra_Index_Response_Rate', 0), 1)
            }
            regions.append(region_metrics)

        # Global 평균 데이터 가져오기
        global_data = data_cache.get_logic_global_data(month)
        global_avg = None
        if global_data:
            global_avg = {
                'region': 'Global',
                'total_score': round(global_data.get('Score', 0), 1),  # 캐시된 Global Score 평균
                'course_completion_rate': round(global_data.get('Course_Completion_Rate', 0), 1),
                'hours_completion_rate': round(global_data.get('Hours_Completion_Rate', 0), 1),
                'new_hire_completion_rate': round(global_data.get('New_Hire_Completion_Rate', 0), 1),
                'eip_completion_rate': round(global_data.get('EIP_Completion_Rate', 0), 1),
                'glp_completion_rate': round(global_data.get('GLP_Completion_Rate', 0), 1),
                'new_leader_rate': round(global_data.get('New_Leader_Completion_Rate', 0), 1),
                'lms_education_rate': round(global_data.get('New_LMS_Course_Rate', 0), 1),
                'lms_mission_rate': round(global_data.get('LMS_Mission_Rate', 0), 1),
                'annual_plan': round(global_data.get('Annual_Plan_Setup_Rate', 0), 1),
                'jam_community': round(global_data.get('JAM_Member_Rate', 0), 1),
                'global_council': round(global_data.get('Global_LD_Council_Rate', 0), 1),
                'infra_index': round(global_data.get('Infra_Index_Response_Rate', 0), 1)
            }

        response_data = {
            'success': True,
            'regions': regions,
            'global_avg': global_avg
        }

        print(f"Debug - 응답 데이터 준비 완료: {len(regions)} 개 Region + Global 평균")
        return jsonify(response_data)

    except Exception as e:
        print(f"Debug - Global region metrics API 에러: {str(e)}")
        import traceback
        print(f"Debug - 에러 상세: {traceback.format_exc()}")
        return jsonify({
            'success': False,
            'error': f'서버 오류가 발생했습니다: {str(e)}'
        }), 500


@blueprint.route('/api/region-infos/<region>/<int:month>')
@login_required
def get_region_infos(region, month):
    """특정 지역의 hr_index_final 데이터를 반환"""
    try:
        import pandas as pd
        from pathlib import Path
        from urllib.parse import unquote

        # URL 디코딩
        region = unquote(region)
        print(f"Debug - Region infos request: {region}, month: {month}")

        # 파일 경로 설정 (절대 경로 사용)
        month_folder = str(month)
        csv_file = DATA_DIR / month_folder / "hr_index_final.csv"

        if not csv_file.exists():
            return jsonify({
                'success': False,
                'error': f'File not found: {csv_file}'
            }), 404

        # CSV 파일 읽기
        df = pd.read_csv(csv_file)
        print(f"Debug - hr_index_final.csv 로드 완료: {len(df)} 행")

        # Final Region 컬럼이 있는지 확인
        if 'Final Region' not in df.columns:
            return jsonify({
                'success': False,
                'error': 'Final Region column not found'
            }), 400

        # 대소문자 구분 없이 region 매칭
        region_data = df[df['Final Region'].str.lower() == region.lower()]

        if region_data.empty:
            return jsonify({
                'success': False,
                'error': f'No data found for region: {region}'
            }), 404

        print(f"Debug - Region 데이터 찾음: {len(region_data)} 행")

        # 지역별 집계 계산
        total_count = len(region_data)
        new_hire_count = len(region_data[region_data['New Hire'] == 'Y']) if 'New Hire' in region_data.columns else 0
        eip_count = len(region_data[region_data['HIPO Type'] == 'EIP']) if 'HIPO Type' in region_data.columns else 0
        glp_count = len(region_data[region_data['HIPO Type'] == 'GLP']) if 'HIPO Type' in region_data.columns else 0
        staff_count = len(region_data[region_data['Staff/Operator'] == 'Staff']) if 'Staff/Operator' in region_data.columns else 0
        operator_count = len(region_data[region_data['Staff/Operator'] == 'Operator']) if 'Staff/Operator' in region_data.columns else 0

        # Region 정보 (첫 번째 행에서 가져오기)
        region_name = region_data['Final Region'].iloc[0] if not region_data.empty else region

        response_data = {
            'success': True,
            'region': region,
            'month': month,
            'data': {
                'region_name': region_name,
                'total_count': total_count,
                'new_hire_count': new_hire_count,
                'eip_count': eip_count,
                'glp_count': glp_count,
                'staff_count': staff_count,
                'operator_count': operator_count
            }
        }

        print(f"Debug - Region infos 응답: {response_data}")
        return jsonify(response_data)

    except Exception as e:
        print(f"Error getting region infos: {e}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@blueprint.route('/api/global-infos/<int:month>')
# @login_required  # 임시로 주석 처리
def get_global_infos(month):
    """Global hr_index_final 데이터를 반환"""
    try:
        import pandas as pd
        from pathlib import Path

        print(f"Debug - Global infos request: month: {month}")
        print(f"Debug - DATA_DIR: {DATA_DIR}")

        # 파일 경로 설정 (절대 경로 사용)
        month_folder = str(month)
        csv_file = DATA_DIR / month_folder / "hr_index_final.csv"
        print(f"Debug - csv_file: {csv_file}")
        print(f"Debug - csv_file.exists(): {csv_file.exists()}")

        if not csv_file.exists():
            return jsonify({
                'success': False,
                'error': f'File not found: {csv_file}'
            }), 404

        # CSV 파일 읽기
        df = pd.read_csv(csv_file)
        print(f"Debug - hr_index_final.csv 로드 완료: {len(df)} 행")

        # 전체 데이터 집계 (Global)
        total_count = len(df)
        new_hire_count = len(df[df['New Hire'] == 'Y']) if 'New Hire' in df.columns else 0
        eip_count = len(df[df['HIPO Type'] == 'EIP']) if 'HIPO Type' in df.columns else 0
        glp_count = len(df[df['HIPO Type'] == 'GLP']) if 'HIPO Type' in df.columns else 0
        staff_count = len(df[df['Staff/Operator'] == 'Staff']) if 'Staff/Operator' in df.columns else 0
        operator_count = len(df[df['Staff/Operator'] == 'Operator']) if 'Staff/Operator' in df.columns else 0

        response_data = {
            'success': True,
            'month': month,
            'data': {
                'total_count': total_count,
                'new_hire_count': new_hire_count,
                'eip_count': eip_count,
                'glp_count': glp_count,
                'staff_count': staff_count,
                'operator_count': operator_count
            }
        }

        print(f"Debug - Global infos 응답: {response_data}")
        return jsonify(response_data)

    except Exception as e:
        print(f"Error getting global infos: {e}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@blueprint.route('/api/region-course-list/<region>/<int:month>')
# @login_required  # 임시로 주석 처리
def get_region_course_list(region, month):
    """특정 지역의 완료된 과정 리스트를 반환"""
    try:
        import pandas as pd
        from urllib.parse import unquote
        import os

        # URL 디코딩
        region = unquote(region)
        print(f"Debug - 지역 과정리스트 요청: region={region}, month={month}")

        # join_hr_lms.csv 파일 경로
        csv_path = str(DATA_DIR / f"{month}" / "join_hr_lms.csv")

        if not os.path.exists(csv_path):
            print(f"Debug - 파일 없음: {csv_path}")
            return jsonify({
                'success': False,
                'error': f'{month}월 데이터가 존재하지 않습니다.'
            }), 404

        # CSV 파일 읽기
        df = pd.read_csv(csv_path)
        print(f"Debug - join_hr_lms.csv 로드 완료: {len(df)} 행")
        print(f"Debug - CSV 컬럼 목록: {list(df.columns)}")

        # Final Region 컬럼이 있는지 확인
        if 'Final Region' not in df.columns:
            return jsonify({
                'success': False,
                'error': 'Final Region 컬럼을 찾을 수 없습니다.'
            }), 400

        # Final Region 고유값들 확인
        unique_regions = df['Final Region'].unique()
        print(f"Debug - Final Region 고유값들: {list(unique_regions)[:10]}...")  # 처음 10개만 출력
        print(f"Debug - 요청된 region: '{region}'")

        # 해당 region 데이터 필터링 (대소문자 구분 없이)
        region_data = df[df['Final Region'].str.lower() == region.lower()]
        print(f"Debug - 대소문자 변환 후 매칭 시도: '{region.lower()}'")
        print(f"Debug - {region} 지역 데이터 필터링 완료: {len(region_data)} 행")

        if region_data.empty:
            # 정확한 매칭이 안될 경우 부분 매칭 시도
            print(f"Debug - 정확한 매칭 실패, 부분 매칭 시도...")
            partial_matches = df[df['Final Region'].str.lower().str.contains(region.lower(), na=False)]
            print(f"Debug - 부분 매칭 결과: {len(partial_matches)} 행")
            if len(partial_matches) > 0:
                print(f"Debug - 부분 매칭된 Final Region 값들: {partial_matches['Final Region'].unique()}")

            return jsonify({
                'success': False,
                'error': f'{region} 지역의 데이터를 찾을 수 없습니다. 사용 가능한 지역: {list(unique_regions)[:20]}...'
            }), 404

        # 완료된 과정만 필터링 (-C로 끝나는 것)
        completed_data = region_data[region_data['Completion status'].str.endswith('-C', na=False)]

        if completed_data.empty:
            return jsonify({
                'success': True,
                'courses': [],
                'staff_unique_count': 0,
                'message': f'{region} 지역에 완료된 과정이 없습니다.'
            })

        print(f"Debug - 완료된 과정 필터링: {len(completed_data)} 행")

        # 기준월과 일치하는 완료일 필터링
        # Completion Date가 비어있거나 잘못된 형식인 경우를 처리
        print(f"Debug - Completion Date 샘플: {completed_data['Completion Date'].head().tolist()}")

        # Completion Date가 비어있지 않은 경우만 처리
        valid_dates = completed_data['Completion Date'].notna() & (completed_data['Completion Date'] != '') & (completed_data['Completion Date'] != ' ')
        completed_data_valid = completed_data[valid_dates].copy()

        if completed_data_valid.empty:
            print(f"Debug - 유효한 Completion Date가 없음")
            return jsonify({
                'success': True,
                'courses': [],
                'staff_unique_count': 0,
                'message': f'{region} 지역에 유효한 완료일이 있는 과정이 없습니다.'
            })

        # Completion Date가 float 형태(20250812.0)인 경우를 처리
        completed_data_valid['Completion Date'] = completed_data_valid['Completion Date'].astype(str).str.replace('.0', '')
        completed_data_valid['Completion Date'] = pd.to_datetime(completed_data_valid['Completion Date'], format='%Y%m%d', errors='coerce')

        # 날짜 변환이 실패한 경우 제외
        valid_dates_parsed = completed_data_valid['Completion Date'].notna()
        completed_data_valid = completed_data_valid[valid_dates_parsed]

        if completed_data_valid.empty:
            print(f"Debug - 날짜 파싱 후 유효한 데이터가 없음")
            return jsonify({
                'success': True,
                'courses': [],
                'staff_unique_count': 0,
                'message': f'{region} 지역에 유효한 날짜 형식의 완료 과정이 없습니다.'
            })

        month_data = completed_data_valid[completed_data_valid['Completion Date'].dt.month == month]

        if month_data.empty:
            return jsonify({
                'success': True,
                'courses': [],
                'staff_unique_count': 0,
                'message': f'{region} 지역에 {month}월 완료된 과정이 없습니다.'
            })

        print(f"Debug - {month}월 완료된 과정: {len(month_data)} 행")

        # 그룹핑하여 과정별 집계
        course_summary = month_data.groupby('Course name').agg({
            'category_1': 'first',   # 카테고리 대 (첫 번째 값 사용)
            'category_2': 'first',   # 카테고리 중 (첫 번째 값 사용)
            'Category': 'first',     # 카테고리 소 (첫 번째 값 사용)
            'Final Region': 'count', # 이수인원
            'Education Hours': 'sum' # 총 이수시간
        }).rename(columns={
            'Final Region': 'participant_count',
            'Education Hours': 'total_hours'
        })

        # 인덱스를 컬럼으로 변환
        course_summary = course_summary.reset_index()

        print(f"Debug - 과정별 집계 완료: {len(course_summary)} 과정")

        # JSON 직렬화를 위한 안전한 값 변환 함수
        def safe_value(value):
            if pd.isna(value):
                return None
            return value

        # 과정 리스트 생성
        courses = []
        for _, row in course_summary.iterrows():
            courses.append({
                'course_name': safe_value(row['Course name']),
                'category_large': safe_value(row['category_1']),
                'category_medium': safe_value(row['category_2']),
                'category_small': safe_value(row['Category']),
                'participant_count': int(row['participant_count']),
                'total_hours': float(row['total_hours'])
            })

        print(f"Debug - 첫 번째 과정명: {courses[0]['course_name'] if courses else 'None'}")

        # Staff 고유 인원 수 계산: hr_index_final.csv 에서 Staff/Operator == 'Staff' 이고 Employee Number 고유값
        staff_unique_count = None
        try:
            hr_idx_path = str(DATA_DIR / f"{month}" / "hr_index_final.csv")
            if os.path.exists(hr_idx_path):
                hr_df = pd.read_csv(hr_idx_path)
                # 컬럼 이름 표준화 시도
                emp_col = None
                for cand in ['Emp. No.', 'Employee Number', 'Employee_Number', 'EmployeeNumber']:
                    if cand in hr_df.columns:
                        emp_col = cand
                        break

                if emp_col:
                    # Final Region으로 필터링하고 Staff만 선택
                    region_staff = hr_df[
                        (hr_df['Final Region'].str.lower() == region.lower()) &
                        (hr_df['Staff/Operator'] == 'Staff')
                    ]
                    staff_unique_count = region_staff[emp_col].nunique()
                    print(f"Debug - {region} 지역 Staff 고유 인원: {staff_unique_count}명")
                else:
                    print(f"Debug - Employee Number 컬럼을 찾을 수 없음. 사용 가능한 컬럼: {list(hr_df.columns)[:10]}...")
                    staff_unique_count = 0
            else:
                print(f"Debug - hr_index_final.csv 파일 없음")
                staff_unique_count = 0
        except Exception as e:
            print(f"Debug - Staff 고유 인원 계산 오류: {e}")
            staff_unique_count = 0

        response_data = {
            'success': True,
            'courses': courses,
            'staff_unique_count': staff_unique_count or 0
        }

        print(f"Debug - 응답 데이터 준비 완료: {len(courses)} 과정, Staff {staff_unique_count}명")
        return jsonify(response_data)

    except Exception as e:
        print(f"Debug - Region course list API 에러: {str(e)}")
        import traceback
        print(f"Debug - 에러 상세: {traceback.format_exc()}")
        traceback.print_exc()
        return jsonify({
            'success': False,
            'error': f'서버 오류가 발생했습니다: {str(e)}'
        }), 500


@blueprint.route('/api/region-subsidiary-list/<region>/<int:month>')
# @login_required  # 임시로 주석 처리
def get_region_subsidiary_list(region, month):
    """특정 지역의 Subsidiary 리스트를 반환"""
    try:
        import pandas as pd
        from urllib.parse import unquote
        import os

        # URL 디코딩
        region = unquote(region)
        print(f"Debug - 지역 Subsidiary 리스트 요청: region={region}, month={month}")

        # hr_index_final.csv 파일 경로
        csv_path = str(DATA_DIR / f"{month}" / "hr_index_final.csv")

        if not os.path.exists(csv_path):
            print(f"Debug - 파일 없음: {csv_path}")
            return jsonify({
                'success': False,
                'error': f'{month}월 데이터가 존재하지 않습니다.'
            }), 404

        # CSV 파일 읽기
        df = pd.read_csv(csv_path)
        print(f"Debug - hr_index_final.csv 로드 완료: {len(df)} 행")

        # Final Region 컬럼이 있는지 확인
        if 'Final Region' not in df.columns:
            return jsonify({
                'success': False,
                'error': 'Final Region 컬럼을 찾을 수 없습니다.'
            }), 400

        # 해당 region 데이터 필터링 (대소문자 구분 없이)
        region_data = df[df['Final Region'].str.lower() == region.lower()]
        print(f"Debug - {region} 지역 데이터 필터링 완료: {len(region_data)} 행")

        if region_data.empty:
            return jsonify({
                'success': False,
                'error': f'{region} 지역의 데이터를 찾을 수 없습니다.'
            }), 404

        # Final Sub.로 그룹핑하여 집계
        subsidiary_summary = region_data.groupby('Final Sub.').agg({
            'Final Sub.': 'count',  # Total Count
            'New Hire': lambda x: (x == 'Y').sum(),  # New Hire Count
            'HIPO Type': lambda x: (x == 'EIP').sum(),  # EIP Count
            'HIPO Type': lambda x: (x == 'GLP').sum(),  # GLP Count (이건 별도로 계산)
            'Staff/Operator': lambda x: (x == 'Staff').sum(),  # Staff Count
            'Staff/Operator': lambda x: (x == 'Operator').sum()  # Operator Count (이것도 별도로 계산)
        }).rename(columns={'Final Sub.': 'total_count'})

        # HIPO Type과 Staff/Operator는 별도로 계산해야 함
        subsidiary_summary = region_data.groupby('Final Sub.').agg({
            'Final Sub.': 'count',  # Total Count
        }).rename(columns={'Final Sub.': 'total_count'})

        # New Hire Count
        new_hire_counts = region_data.groupby('Final Sub.')['New Hire'].apply(lambda x: (x == 'Y').sum())
        subsidiary_summary['new_hire_count'] = new_hire_counts

        # EIP Count
        eip_counts = region_data.groupby('Final Sub.')['HIPO Type'].apply(lambda x: (x == 'EIP').sum())
        subsidiary_summary['eip_count'] = eip_counts

        # GLP Count
        glp_counts = region_data.groupby('Final Sub.')['HIPO Type'].apply(lambda x: (x == 'GLP').sum())
        subsidiary_summary['glp_count'] = glp_counts

        # Staff Count
        staff_counts = region_data.groupby('Final Sub.')['Staff/Operator'].apply(lambda x: (x == 'Staff').sum())
        subsidiary_summary['staff_count'] = staff_counts

        # Operator Count
        operator_counts = region_data.groupby('Final Sub.')['Staff/Operator'].apply(lambda x: (x == 'Operator').sum())
        subsidiary_summary['operator_count'] = operator_counts

        # 인덱스를 컬럼으로 변환
        subsidiary_summary = subsidiary_summary.reset_index()

        print(f"Debug - Subsidiary 집계 완료: {len(subsidiary_summary)} 개")

        # JSON 직렬화를 위한 안전한 값 변환 함수
        def safe_value(value):
            if pd.isna(value):
                return None
            return value

        subsidiaries = []
        for _, row in subsidiary_summary.iterrows():
            subsidiaries.append({
                'final_sub': safe_value(row['Final Sub.']),
                'total_count': int(row['total_count']),
                'new_hire_count': int(row['new_hire_count']),
                'eip_count': int(row['eip_count']),
                'glp_count': int(row['glp_count']),
                'staff_count': int(row['staff_count']),
                'operator_count': int(row['operator_count'])
            })

        response_data = {
            'success': True,
            'subsidiaries': subsidiaries
        }

        print(f"Debug - 응답 데이터 준비 완료: {len(subsidiaries)} 개 Subsidiary")
        return jsonify(response_data)

    except Exception as e:
        print(f"Debug - Region subsidiary list API 에러: {str(e)}")
        import traceback
        print(f"Debug - 에러 상세: {traceback.format_exc()}")
        return jsonify({
            'success': False,
            'error': f'서버 오류가 발생했습니다: {str(e)}'
        }), 500


@blueprint.route('/api/region-subsidiary-metrics/<region>/<int:month>')
@login_required
def get_region_subsidiary_metrics(region, month):
    """특정 지역의 Subsidiary별 지표 데이터를 반환"""
    try:
        import pandas as pd
        from urllib.parse import unquote
        import os

        # URL 디코딩
        region = unquote(region)
        print(f"Debug - 지역 Subsidiary 지표 요청: region={region}, month={month}")

        # logic.csv 파일 경로
        csv_path = str(DATA_DIR / f"{month}" / "logic.csv")

        if not os.path.exists(csv_path):
            print(f"Debug - 파일 없음: {csv_path}")
            return jsonify({
                'success': False,
                'error': f'{month}월 logic.csv 데이터가 존재하지 않습니다.'
            }), 404

        # CSV 파일 읽기
        df = pd.read_csv(csv_path)
        print(f"Debug - CSV 파일 로드 완료: {len(df)} 행")

        # 컬럼명 확인
        if 'Final Region' not in df.columns:
            return jsonify({
                'success': False,
                'error': 'Final Region 컬럼을 찾을 수 없습니다.'
            }), 400

        # NaN 값을 안전하게 처리하는 함수
        def safe_round(value, decimals=1):
            """NaN을 0으로 처리하고 round"""
            if pd.isna(value):
                return 0
            return round(float(value), decimals)

        # 해당 region 데이터 필터링 (대소문자 구분 없이)
        region_data = df[df['Final Region'].str.lower() == region.lower()]
        print(f"Debug - {region} 지역 데이터 필터링 완료: {len(region_data)} 행")

        if region_data.empty:
            return jsonify({
                'success': False,
                'error': f'{region} 지역의 데이터를 찾을 수 없습니다.'
            }), 404

        # Subsidiary별 지표 데이터 구성
        subsidiaries = []
        for _, row in region_data.iterrows():
            # Y/N 값 가져오기
            new_lms_course = row.get('New LMS Course', 'N')
            lms_mission = row.get('LMS Mission', 'N')
            annual_plan = row.get('Annual Plan Setup', 'N')
            jam_member = row.get('JAM Member', 'N')
            global_ld_council = row.get('Global L&D Council', 'N')
            infra_index = row.get('Infra index response', 'N')

            subsidiary_data = {
                'subsidiary': row.get('Subsidiary', '-'),
                'total_score': safe_round(row.get('Score', 0)),  # logic.csv의 Score 값
                'course_completion_rate': safe_round(row.get('Course_Completion_Rate', 0)),
                'hours_completion_rate': safe_round(row.get('Hours_Completion_Rate', 0)),
                'new_hire_completion_rate': safe_round(row.get('New_Hire_Completion_Rate', 0)),
                'eip_completion_rate': safe_round(row.get('EIP_Completion_Rate', 0)),
                'glp_completion_rate': safe_round(row.get('GLP_Completion_Rate', 0)),
                'new_leader_rate': safe_round(row.get('New_Leader_Completion_Rate', 0)),
                'lms_education_rate': new_lms_course,
                'lms_mission_rate': lms_mission,
                'annual_plan': annual_plan,
                'jam_community': jam_member,
                'global_council': global_ld_council,
                'infra_index': infra_index
            }
            subsidiaries.append(subsidiary_data)

        # Region 평균 데이터 가져오기 (캐시에서)
        region_avg = data_cache.get_logic_region_data(month, region)
        region_avg_data = None

        if region_avg:
            region_avg_data = {
                'subsidiary': 'Region',
                'total_score': safe_round(region_avg.get('Score', 0)),  # 캐시된 Region Score 평균
                'course_completion_rate': safe_round(region_avg.get('Course_Completion_Rate', 0)),
                'hours_completion_rate': safe_round(region_avg.get('Hours_Completion_Rate', 0)),
                'new_hire_completion_rate': safe_round(region_avg.get('New_Hire_Completion_Rate', 0)),
                'eip_completion_rate': safe_round(region_avg.get('EIP_Completion_Rate', 0)),
                'glp_completion_rate': safe_round(region_avg.get('GLP_Completion_Rate', 0)),
                'new_leader_rate': safe_round(region_avg.get('New_Leader_Completion_Rate', 0)),
                'lms_education_rate': safe_round(region_avg.get('New_LMS_Course_Rate', 0)),
                'lms_mission_rate': safe_round(region_avg.get('LMS_Mission_Rate', 0)),
                'annual_plan': safe_round(region_avg.get('Annual_Plan_Setup_Rate', 0)),
                'jam_community': safe_round(region_avg.get('JAM_Member_Rate', 0)),
                'global_council': safe_round(region_avg.get('Global_LD_Council_Rate', 0)),
                'infra_index': safe_round(region_avg.get('Infra_Index_Response_Rate', 0))
            }

        response_data = {
            'success': True,
            'subsidiaries': subsidiaries,
            'region_avg': region_avg_data
        }

        print(f"Debug - 응답 데이터 준비 완료: {len(subsidiaries)} 개 Subsidiary 지표 + Region 평균")
        return jsonify(response_data)

    except Exception as e:
        print(f"Debug - Region subsidiary metrics API 에러: {str(e)}")
        import traceback
        print(f"Debug - 에러 상세: {traceback.format_exc()}")
        return jsonify({
            'success': False,
            'error': f'서버 오류가 발생했습니다: {str(e)}'
        }), 500


@blueprint.route('/api/course-list/<subsidiary>/<int:month>')
@login_required
def get_course_list(subsidiary, month):
    """특정 법인의 완료된 과정 리스트를 반환"""
    try:
        import pandas as pd
        import os

        print(f"Debug - 과정리스트 요청: subsidiary={subsidiary}, month={month}")

        # join_hr_lms.csv 파일 경로
        csv_path = str(DATA_DIR / f"{month}" / "join_hr_lms.csv")

        if not os.path.exists(csv_path):
            print(f"Debug - 파일 없음: {csv_path}")
            return jsonify({
                'success': False,
                'error': f'{month}월 데이터를 찾을 수 없습니다.'
            }), 404

        # CSV 파일 읽기
        df = pd.read_csv(csv_path)
        print(f"Debug - CSV 파일 로드 완료: {len(df)} 행")

        # 1) Final Sub.이 선택된 subsidiary와 같은 것만 1차 필터링 (케이스 무시)
        df_sub = df[df['Final Sub.'].astype(str).str.lower() == str(subsidiary).lower()].copy()
        print(f"Debug - 1차 필터링(법인) 후: {len(df_sub)} 행")

        # 2) Completion status가 '-C'로 끝나는 것만 필터링
        df_sub_completed = df_sub[df_sub['Completion status'].astype(str).str.endswith('-C', na=False)].copy()
        print(f"Debug - 2차 필터링(완료) 후: {len(df_sub_completed)} 행")

        # 3) Completion Date에서 월 추출 (yyyymmdd 형태)
        df_sub_completed['completion_month'] = pd.to_datetime(df_sub_completed['Completion Date'], format='%Y%m%d', errors='coerce').dt.month

        # 4) 기준월과 같은 것만 필터링
        monthly_completed = df_sub_completed[df_sub_completed['completion_month'] == month]
        print(f"Debug - 3차 필터링(기준월 {month}) 후: {len(monthly_completed)} 행")

        # 최종 데이터프레임
        subsidiary_data = monthly_completed
        print(f"Debug - 최종 대상 행수: {len(subsidiary_data)}")

        if subsidiary_data.empty:
            print(f"Debug - {subsidiary} 법인의 완료된 과정이 없음")
            return jsonify({
                'success': True,
                'courses': [],
                'message': '해당 법인의 완료된 과정이 없습니다.'
            })

        # 그룹핑하여 과정별 집계
        course_summary = (
            subsidiary_data
            .groupby('Course name', dropna=False)
            .agg({
                'category_1': 'max',     # 카테고리 대
                'category_2': 'max',     # 카테고리 중
                'Category': 'max',       # 카테고리 소
                'Final Sub.': 'count',   # 이수인원
                'Education Hours': 'sum' # 총 이수시간
            })
            .rename(columns={
                'Final Sub.': 'participant_count',
                'Education Hours': 'total_hours'
            })
            .reset_index()
        )
        print(f"Debug - course_summary 컬럼: {course_summary.columns.tolist()}")
        print(f"Debug - 첫 번째 행의 Course name: {course_summary.iloc[0]['Course name'] if len(course_summary) > 0 else 'None'}")

        # 결과를 리스트로 변환
        courses = []
        for _, row in course_summary.iterrows():
            # NaN 값을 None으로 변환
            def safe_value(value):
                if pd.isna(value):
                    return None
                return value

            courses.append({
                'course_name': safe_value(row['Course name']),  # Course name 컬럼에서 직접 가져오기
                'category_large': safe_value(row['category_1']),
                'category_medium': safe_value(row['category_2']),
                'category_small': safe_value(row['Category']),
                'participant_count': int(row['participant_count']),
                'total_hours': float(row['total_hours'])
            })

        print(f"Debug - 최종 과정 수: {len(courses)}")
        if courses:
            print(f"Debug - 첫 번째 과정명: {courses[0]['course_name']}")

        # Staff 고유 인원 수 계산: hr_index_final.csv 에서 Staff/Operator == 'Staff' 이고 Employee Number 고유값
        staff_unique_count = None
        try:
            hr_idx_path = str(DATA_DIR / f"{month}" / "hr_index_final.csv")
            if os.path.exists(hr_idx_path):
                hr_df = pd.read_csv(hr_idx_path)
                # 컬럼 이름 표준화 시도
                emp_col = None
                for cand in ['Emp. No.', 'Employee Number', 'Employee_Number', 'EmployeeNumber']:
                    if cand in hr_df.columns:
                        emp_col = cand
                        break
                staffop_col = None
                for cand in ['Staff/Operator', 'Staff_Operator', 'StaffOperator']:
                    if cand in hr_df.columns:
                        staffop_col = cand
                        break
                sub_col = None
                for cand in ['Final Sub.', 'Final Sub', 'Final_Sub']:
                    if cand in hr_df.columns:
                        sub_col = cand
                        break

                if emp_col and staffop_col and sub_col:
                    hr_sub = hr_df[hr_df[sub_col].astype(str).str.lower() == str(subsidiary).lower()].copy()
                    hr_sub_staff = hr_sub[hr_sub[staffop_col].astype(str).str.strip().str.lower() == 'staff']
                    staff_unique_count = int(hr_sub_staff[emp_col].nunique())
                    print(f"Debug - Staff 고유 인원 수: {staff_unique_count}")
                else:
                    print("Debug - hr_index_final.csv 필수 컬럼 누락")
            else:
                print(f"Debug - hr_index_final.csv 없음: {hr_idx_path}")
        except Exception as e2:
            print(f"Debug - Staff 고유 인원 계산 오류: {e2}")

        return jsonify({
            'success': True,
            'courses': courses,
            'total_courses': len(courses),
            'staff_unique_count': staff_unique_count
        })

    except Exception as e:
        import traceback
        print(f"Debug - 과정리스트 오류: {str(e)}")
        print(f"Debug - 상세 오류: {traceback.format_exc()}")
        return jsonify({
            'success': False,
            'error': f'과정리스트를 불러오는 중 오류가 발생했습니다: {str(e)}'
        }), 500


@blueprint.route('/api/region-summary/<region>/<int:month>')
@login_required
def get_region_summary(region, month):
    """특정 지역의 logic.csv 요약 데이터를 반환"""
    try:
        import pandas as pd
        import os

        print(f"Debug - 지역 요약 요청: region={region}, month={month}")

        # logic.csv 파일 경로
        csv_path = str(DATA_DIR / f"{month}" / "logic.csv")

        if not os.path.exists(csv_path):
            print(f"Debug - 파일 없음: {csv_path}")
            return jsonify({
                'success': False,
                'error': f'{month}월 데이터를 찾을 수 없습니다.'
            }), 404

        # CSV 파일 읽기
        df = pd.read_csv(csv_path)
        print(f"Debug - CSV 파일 로드 완료: {len(df)} 행")

        # Final Region이 선택된 region과 같은 것만 필터링
        region_data = df[df['Final Region'].astype(str).str.lower() == str(region).lower()]
        print(f"Debug - {region} 지역 데이터: {len(region_data)} 행")

        if region_data.empty:
            print(f"Debug - {region} 지역 데이터 없음")
            return jsonify({
                'success': True,
                'data': {},
                'message': f'{region} 지역의 데이터가 없습니다.'
            })

        # Final Region으로 그룹핑하고 sum 계산
        # 필요한 컬럼들 정의
        sum_columns = [
            'Planned_Courses', 'Completed_Courses', 'Planned_Hours', 'Actual_Hours',
            'New_Hire_Completed', 'New_Hire_Not_Completed', 'New_Hire_Pending', 'New_Hire_Total',
            'EIP_Completed', 'EIP_Not_Completed', 'EIP_Total',
            'GLP_Completed', 'GLP_Not_Completed', 'GLP_Total'
        ]

        # 존재하는 컬럼만 선택
        available_columns = [col for col in sum_columns if col in region_data.columns]

        # 그룹핑 및 sum 계산
        summary = region_data.groupby('Final Region')[available_columns].sum()

        # 결과를 딕셔너리로 변환
        result = {}
        for col in available_columns:
            if col in summary.columns:
                result[col] = float(summary.iloc[0][col])

        # 비율 계산
        # Course_Completion_Rate = Completed_Courses / Planned_Courses * 100
        if 'Completed_Courses' in result and 'Planned_Courses' in result and result['Planned_Courses'] > 0:
            result['Course_Completion_Rate'] = round(result['Completed_Courses'] / result['Planned_Courses'] * 100, 1)

        # Hours_Completion_Rate = Actual_Hours / Planned_Hours * 100
        if 'Actual_Hours' in result and 'Planned_Hours' in result and result['Planned_Hours'] > 0:
            result['Hours_Completion_Rate'] = round(result['Actual_Hours'] / result['Planned_Hours'] * 100, 1)

        # New_Hire_Completion_Rate = (New_Hire_Completed + New_Hire_Pending) / New_Hire_Total * 100
        if all(col in result for col in ['New_Hire_Completed', 'New_Hire_Pending', 'New_Hire_Total']) and result['New_Hire_Total'] > 0:
            result['New_Hire_Completion_Rate'] = round((result['New_Hire_Completed'] + result['New_Hire_Pending']) / result['New_Hire_Total'] * 100, 1)

        # EIP_Completion_Rate = EIP_Completed / EIP_Total * 100
        if 'EIP_Completed' in result and 'EIP_Total' in result and result['EIP_Total'] > 0:
            result['EIP_Completion_Rate'] = round(result['EIP_Completed'] / result['EIP_Total'] * 100, 1)

        # GLP_Completion_Rate = GLP_Completed / GLP_Total * 100
        if 'GLP_Completed' in result and 'GLP_Total' in result and result['GLP_Total'] > 0:
            result['GLP_Completion_Rate'] = round(result['GLP_Completed'] / result['GLP_Total'] * 100, 1)

        print(f"Debug - 계산된 평균 데이터: {result}")

        return jsonify({
            'success': True,
            'data': result,
            'region': region,
            'month': month,
            'total_records': len(region_data)
        })

    except Exception as e:
        import traceback
        print(f"Debug - 지역 요약 오류: {str(e)}")
        print(f"Debug - 상세 오류: {traceback.format_exc()}")
        return jsonify({
            'success': False,
            'error': f'지역 요약을 불러오는 중 오류가 발생했습니다: {str(e)}'
        }), 500


@blueprint.route('/api/logic-course-completion/<subsidiary>/<int:month>')
@login_required
def get_logic_course_completion(subsidiary, month):
    """logic.csv 에서 특정 법인의 Course_Completion_Rate 를 반환하는 API"""
    try:
        import pandas as pd
        from pathlib import Path

        month_folder = str(month)
        csv_file = DATA_DIR / month_folder / "logic.csv"

        if not csv_file.exists():
            return jsonify({
                'success': False,
                'error': f'{month_folder} logic.csv 파일이 존재하지 않습니다.'
            }), 404

        df = pd.read_csv(csv_file)

        # 디버깅: 컬럼명 출력
        print(f"Debug - Available columns: {list(df.columns)}")
        print(f"Debug - Looking for subsidiary: {subsidiary}")

        # 법인 컬럼 찾기 (대소문자 무관)
        sub_col = None
        for col in df.columns:
            if col.lower() in ['subsidiary', 'final sub.', 'final_sub', 'final_sub.']:
                sub_col = col
                break

        if sub_col is None:
            return jsonify({
                'success': False,
                'error': f'logic.csv 에서 법인 식별 컬럼을 찾을 수 없습니다. 사용 가능한 컬럼: {list(df.columns)}'
            }), 400

        # Course_Completion_Rate 컬럼 찾기 (대소문자 무관)
        rate_col = None
        for col in df.columns:
            if col.lower() in ['course_completion_rate', 'course completion rate', 'completion_rate']:
                rate_col = col
                break

        if rate_col is None:
            return jsonify({
                'success': False,
                'error': f'logic.csv 에서 Course_Completion_Rate 컬럼을 찾을 수 없습니다. 사용 가능한 컬럼: {list(df.columns)}'
            }), 400

        print(f"Debug - Using subsidiary column: {sub_col}")
        print(f"Debug - Using rate column: {rate_col}")

        # 해당 subsidiary 데이터 찾기 (대소문자 무관)
        # 양쪽 모두 소문자로 변환해서 매칭
        df_lower = df.copy()
        df_lower[sub_col] = df_lower[sub_col].astype(str).str.lower()
        subsidiary_lower = subsidiary.lower()

        row = df[df_lower[sub_col] == subsidiary_lower]
        if row.empty:
            # 디버깅: 사용 가능한 subsidiary 값들 출력
            available_subs = df[sub_col].unique()[:10]  # 처음 10개만
            return jsonify({
                'success': False,
                'error': f'{subsidiary} 법인의 logic.csv 데이터가 없습니다. 사용 가능한 법인: {list(available_subs)}'
            }), 404

        value = row[rate_col].iloc[0]
        print(f"Debug - Found value: {value} (type: {type(value)})")

        # Planned_Courses와 Completed_Courses 값 가져오기
        planned_courses = '-'
        completed_courses = '-'

        # Planned_Courses 컬럼 찾기
        planned_col = None
        for col in df.columns:
            if col.lower() in ['planned_courses', 'planned courses', 'planned']:
                planned_col = col
                break

        # Completed_Courses 컬럼 찾기
        completed_col = None
        for col in df.columns:
            if col.lower() in ['completed_courses', 'completed courses', 'completed']:
                completed_col = col
                break

        if planned_col and planned_col in row.columns:
            planned_courses = row[planned_col].iloc[0]
            if pd.isna(planned_courses):
                planned_courses = '-'

        if completed_col and completed_col in row.columns:
            completed_courses = row[completed_col].iloc[0]
            if pd.isna(completed_courses):
                completed_courses = '-'

        # Hours_Completion_Rate, Planned_Hours, Actual_Hours 값 가져오기
        hours_completion_rate = '-'
        planned_hours = '-'
        actual_hours = '-'

        # Hours_Completion_Rate 컬럼 찾기
        hours_rate_col = None
        for col in df.columns:
            if col.lower() in ['hours_completion_rate', 'hours completion rate', 'completion_rate']:
                hours_rate_col = col
                break

        # Planned_Hours 컬럼 찾기
        planned_hours_col = None
        for col in df.columns:
            if col.lower() in ['planned_hours', 'planned hours', 'planned']:
                planned_hours_col = col
                break

        # Actual_Hours 컬럼 찾기
        actual_hours_col = None
        for col in df.columns:
            if col.lower() in ['actual_hours', 'actual hours', 'actual']:
                actual_hours_col = col
                break

        if hours_rate_col and hours_rate_col in row.columns:
            hours_completion_rate = row[hours_rate_col].iloc[0]
            if pd.isna(hours_completion_rate):
                hours_completion_rate = '-'

        if planned_hours_col and planned_hours_col in row.columns:
            planned_hours = row[planned_hours_col].iloc[0]
            if pd.isna(planned_hours):
                planned_hours = '-'

        if actual_hours_col and actual_hours_col in row.columns:
            actual_hours = row[actual_hours_col].iloc[0]
            if pd.isna(actual_hours):
                actual_hours = '-'

        # New Hire 지표: New_Hire_Completion_Rate, New_Hire_Total, New_Hire_Completed,
        #                New_Hire_Not_Completed, New_Hire_Pending
        nh_completion_rate = '-'
        nh_total = '-'
        nh_completed = '-'
        nh_not_completed = '-'
        nh_pending = '-'

        def find_col(candidates):
            for col in df.columns:
                if col.lower() in candidates:
                    return col
            return None

        nh_rate_col = find_col(['new_hire_completion_rate', 'new hire completion rate'])
        nh_total_col = find_col(['new_hire_total', 'new hire total'])
        nh_completed_col = find_col(['new_hire_completed', 'new hire completed'])
        nh_not_completed_col = find_col(['new_hire_not_completed', 'new hire not completed'])
        nh_pending_col = find_col(['new_hire_pending', 'new hire pending'])

        if nh_rate_col and nh_rate_col in row.columns:
            nh_completion_rate = row[nh_rate_col].iloc[0]
            if pd.isna(nh_completion_rate):
                nh_completion_rate = '-'

        if nh_total_col and nh_total_col in row.columns:
            nh_total = row[nh_total_col].iloc[0]
            if pd.isna(nh_total):
                nh_total = '-'

        if nh_completed_col and nh_completed_col in row.columns:
            nh_completed = row[nh_completed_col].iloc[0]
            if pd.isna(nh_completed):
                nh_completed = '-'

        if nh_not_completed_col and nh_not_completed_col in row.columns:
            nh_not_completed = row[nh_not_completed_col].iloc[0]
            if pd.isna(nh_not_completed):
                nh_not_completed = '-'

        if nh_pending_col and nh_pending_col in row.columns:
            nh_pending = row[nh_pending_col].iloc[0]
            if pd.isna(nh_pending):
                nh_pending = '-'

        # EIP 지표: EIP_Completion_Rate, EIP_Total, EIP_Completed, EIP_Not_Completed
        eip_completion_rate = '-'
        eip_total = '-'
        eip_completed = '-'
        eip_not_completed = '-'

        eip_rate_col = find_col(['eip_completion_rate', 'eip completion rate'])
        eip_total_col = find_col(['eip_total', 'eip total'])
        eip_completed_col = find_col(['eip_completed', 'eip completed'])
        eip_not_completed_col = find_col(['eip_not_completed', 'eip not completed'])

        if eip_rate_col and eip_rate_col in row.columns:
            eip_completion_rate = row[eip_rate_col].iloc[0]
            if pd.isna(eip_completion_rate):
                eip_completion_rate = '-'

        if eip_total_col and eip_total_col in row.columns:
            eip_total = row[eip_total_col].iloc[0]
            if pd.isna(eip_total):
                eip_total = '-'

        if eip_completed_col and eip_completed_col in row.columns:
            eip_completed = row[eip_completed_col].iloc[0]
            if pd.isna(eip_completed):
                eip_completed = '-'

        if eip_not_completed_col and eip_not_completed_col in row.columns:
            eip_not_completed = row[eip_not_completed_col].iloc[0]
            if pd.isna(eip_not_completed):
                eip_not_completed = '-'

        # GLP 지표: GLP_Completion_Rate, GLP_Total, GLP_Completed, GLP_Not_Completed
        glp_completion_rate = '-'
        glp_total = '-'
        glp_completed = '-'
        glp_not_completed = '-'

        glp_rate_col = find_col(['glp_completion_rate', 'glp completion rate'])
        glp_total_col = find_col(['glp_total', 'glp total'])
        glp_completed_col = find_col(['glp_completed', 'glp completed'])
        glp_not_completed_col = find_col(['glp_not_completed', 'glp not completed'])

        if glp_rate_col and glp_rate_col in row.columns:
            glp_completion_rate = row[glp_rate_col].iloc[0]
            if pd.isna(glp_completion_rate):
                glp_completion_rate = '-'

        if glp_total_col and glp_total_col in row.columns:
            glp_total = row[glp_total_col].iloc[0]
            if pd.isna(glp_total):
                glp_total = '-'

        if glp_completed_col and glp_completed_col in row.columns:
            glp_completed = row[glp_completed_col].iloc[0]
            if pd.isna(glp_completed):
                glp_completed = '-'

        if glp_not_completed_col and glp_not_completed_col in row.columns:
            glp_not_completed = row[glp_not_completed_col].iloc[0]
            if pd.isna(glp_not_completed):
                glp_not_completed = '-'

        # New Leader 지표: New_Leader_Completion_Rate, New_Leader_Total, New_Leader_Completed, New_Leader_Not_Completed
        new_leader_completion_rate = '-'
        new_leader_total = '-'
        new_leader_completed = '-'
        new_leader_not_completed = '-'

        new_leader_rate_col = find_col(['new_leader_completion_rate', 'new leader completion rate'])
        new_leader_total_col = find_col(['new_leader_total', 'new leader total'])
        new_leader_completed_col = find_col(['new_leader_completed', 'new leader completed'])
        new_leader_not_completed_col = find_col(['new_leader_not_completed', 'new leader not completed'])

        if new_leader_rate_col and new_leader_rate_col in row.columns:
            new_leader_completion_rate = row[new_leader_rate_col].iloc[0]
            if pd.isna(new_leader_completion_rate):
                new_leader_completion_rate = '-'

        if new_leader_total_col and new_leader_total_col in row.columns:
            new_leader_total = row[new_leader_total_col].iloc[0]
            if pd.isna(new_leader_total):
                new_leader_total = '-'

        if new_leader_completed_col and new_leader_completed_col in row.columns:
            new_leader_completed = row[new_leader_completed_col].iloc[0]
            if pd.isna(new_leader_completed):
                new_leader_completed = '-'

        if new_leader_not_completed_col and new_leader_not_completed_col in row.columns:
            new_leader_not_completed = row[new_leader_not_completed_col].iloc[0]
            if pd.isna(new_leader_not_completed):
                new_leader_not_completed = '-'

        try:
            # 퍼센트 표기 정규화 (0-1 값이면 0-100으로 변환)
            if pd.notna(value):
                if isinstance(value, str) and value.strip().endswith('%'):
                    value_num = float(value.strip().replace('%', ''))
                else:
                    value_num = float(value)
                    if value_num <= 1:
                        value_num = value_num * 100.0
                value = round(value_num, 1)
        except Exception:
            # 숫자 변환 실패 시 문자열 그대로 반환
            pass

        # Global 데이터 가져오기
        global_data = data_cache.get_logic_global_data(month)

        # Region 데이터 가져오기 (subsidiary detail API에서 region 정보 사용)
        region_data = None
        if 'Final Region' in df.columns:
            region = row['Final Region'].iloc[0] if not row.empty else None
            if region and pd.notna(region):
                region_data = data_cache.get_logic_region_data(month, region)
                print(f"Debug - Region 데이터 로드: {region}, 데이터 존재: {region_data is not None}")
                # region_data가 None이면 빈 딕셔너리로 설정
                if region_data is None:
                    region_data = {}

        # global_data가 None이면 빈 딕셔너리로 설정
        if global_data is None:
            global_data = {}

        # Score 값 가져오기
        score = '-'
        score_col = find_col(['score'])
        if score_col and score_col in row.columns:
            score_val = row[score_col].iloc[0]
            if pd.notna(score_val):
                score = round(float(score_val), 1)

        # Index Management에서 추가된 6개 컬럼 값 가져오기
        new_lms_course = '-'
        lms_mission = '-'
        annual_plan_setup = '-'
        jam_member = '-'
        global_ld_council = '-'
        infra_index_response = '-'

        # 컬럼명 매핑 (logic.csv의 컬럼명과 매칭)
        new_lms_col = find_col(['new lms course', 'new_lms_course'])
        lms_mission_col = find_col(['lms mission', 'lms_mission'])
        annual_plan_col = find_col(['annual plan setup', 'annual_plan_setup'])
        jam_member_col = find_col(['jam member', 'jam_member'])
        global_ld_col = find_col(['global l&d council', 'global_l&d_council'])
        infra_index_col = find_col(['infra index response', 'infra_index_response'])

        if new_lms_col and new_lms_col in row.columns:
            new_lms_course = row[new_lms_col].iloc[0]
            if pd.isna(new_lms_course):
                new_lms_course = '-'

        if lms_mission_col and lms_mission_col in row.columns:
            lms_mission = row[lms_mission_col].iloc[0]
            if pd.isna(lms_mission):
                lms_mission = '-'

        if annual_plan_col and annual_plan_col in row.columns:
            annual_plan_setup = row[annual_plan_col].iloc[0]
            if pd.isna(annual_plan_setup):
                annual_plan_setup = '-'

        if jam_member_col and jam_member_col in row.columns:
            jam_member = row[jam_member_col].iloc[0]
            if pd.isna(jam_member):
                jam_member = '-'

        if global_ld_col and global_ld_col in row.columns:
            global_ld_council = row[global_ld_col].iloc[0]
            if pd.isna(global_ld_council):
                global_ld_council = '-'

        if infra_index_col and infra_index_col in row.columns:
            infra_index_response = row[infra_index_col].iloc[0]
            if pd.isna(infra_index_response):
                infra_index_response = '-'

        # 모든 값을 JSON 직렬화 가능하도록 변환
        import numpy as np
        def safe_convert(val):
            if pd.isna(val):
                return None
            # numpy 타입을 Python 기본 타입으로 변환
            if isinstance(val, (np.integer, np.int64, np.int32)):
                return int(val)
            if isinstance(val, (np.floating, np.float64, np.float32)):
                return float(val)
            if isinstance(val, (int, float)):
                return val
            return str(val)

        return jsonify({
            'success': True,
            'subsidiary': subsidiary,
            'month': month,
            'score': safe_convert(score),
            'course_completion_rate': safe_convert(value),
            'planned_courses': safe_convert(planned_courses),
            'completed_courses': safe_convert(completed_courses),
            'hours_completion_rate': safe_convert(hours_completion_rate),
            'planned_hours': safe_convert(planned_hours),
            'actual_hours': safe_convert(actual_hours),
            'new_hire_completion_rate': safe_convert(nh_completion_rate),
            'new_hire_total': safe_convert(nh_total),
            'new_hire_completed': safe_convert(nh_completed),
            'new_hire_not_completed': safe_convert(nh_not_completed),
            'new_hire_pending': safe_convert(nh_pending),
            'eip_completion_rate': safe_convert(eip_completion_rate),
            'eip_total': safe_convert(eip_total),
            'eip_completed': safe_convert(eip_completed),
            'eip_not_completed': safe_convert(eip_not_completed),
            'glp_completion_rate': safe_convert(glp_completion_rate),
            'glp_total': safe_convert(glp_total),
            'glp_completed': safe_convert(glp_completed),
            'glp_not_completed': safe_convert(glp_not_completed),
            'new_leader_completion_rate': safe_convert(new_leader_completion_rate),
            'new_leader_total': safe_convert(new_leader_total),
            'new_leader_completed': safe_convert(new_leader_completed),
            'new_leader_not_completed': safe_convert(new_leader_not_completed),
            'new_lms_course': safe_convert(new_lms_course),
            'lms_mission': safe_convert(lms_mission),
            'annual_plan_setup': safe_convert(annual_plan_setup),
            'jam_member': safe_convert(jam_member),
            'global_ld_council': safe_convert(global_ld_council),
            'infra_index_response': safe_convert(infra_index_response),
            'global_data': global_data if global_data else {},
            'region_data': region_data if region_data else {}
        })
    except Exception as e:
        import traceback
        print(f"Debug - logic course completion API 에러: {str(e)}")
        print(f"Debug - 에러 상세: {traceback.format_exc()}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@blueprint.route('/api/all-subsidiary-metrics/<int:month>')
@login_required
def get_all_subsidiary_metrics(month):
    """모든 Subsidiary의 지표 데이터를 반환 (logic.csv 전체)"""
    try:
        import pandas as pd
        import os

        print(f"Debug - 전체 Subsidiary 지표 요청: month={month}")

        # logic.csv 파일 경로
        csv_path = str(DATA_DIR / f"{month}" / "logic.csv")

        if not os.path.exists(csv_path):
            print(f"Debug - 파일 없음: {csv_path}")
            return jsonify({
                'success': False,
                'error': f'{month}월 logic.csv 데이터가 존재하지 않습니다.'
            }), 404

        # CSV 파일 읽기
        df = pd.read_csv(csv_path)
        print(f"Debug - CSV 파일 로드 완료: {len(df)} 행")

        # NaN 값을 안전하게 처리하는 함수
        def safe_round(value, decimals=1):
            """NaN을 0으로 처리하고 round"""
            if pd.isna(value):
                return 0
            return round(float(value), decimals)

        # Subsidiary별 지표 데이터 구성
        subsidiaries = []
        for _, row in df.iterrows():
            # Y/N 값 가져오기
            new_lms_course = row.get('New LMS Course', 'N')
            lms_mission = row.get('LMS Mission', 'N')
            annual_plan = row.get('Annual Plan Setup', 'N')
            jam_member = row.get('JAM Member', 'N')
            global_ld_council = row.get('Global L&D Council', 'N')
            infra_index = row.get('Infra index response', 'N')

            subsidiary_data = {
                'subsidiary': row.get('Subsidiary', '-'),
                'final_region': row.get('Final Region', '-'),
                'total_score': safe_round(row.get('Score', 0)),
                'course_completion_rate': safe_round(row.get('Course_Completion_Rate', 0)),
                'hours_completion_rate': safe_round(row.get('Hours_Completion_Rate', 0)),
                'new_hire_completion_rate': safe_round(row.get('New_Hire_Completion_Rate', 0)),
                'eip_completion_rate': safe_round(row.get('EIP_Completion_Rate', 0)),
                'glp_completion_rate': safe_round(row.get('GLP_Completion_Rate', 0)),
                'new_leader_rate': safe_round(row.get('New_Leader_Completion_Rate', 0)),
                'lms_education_rate': new_lms_course,
                'lms_mission_rate': lms_mission,
                'annual_plan': annual_plan,
                'jam_community': jam_member,
                'global_council': global_ld_council,
                'infra_index': infra_index
            }
            subsidiaries.append(subsidiary_data)

        response_data = {
            'success': True,
            'subsidiaries': subsidiaries
        }

        print(f"Debug - 응답 데이터 준비 완료: {len(subsidiaries)} 개 Subsidiary 지표")
        return jsonify(response_data)

    except Exception as e:
        print(f"Debug - All subsidiary metrics API 에러: {str(e)}")
        import traceback
        print(f"Debug - 에러 상세: {traceback.format_exc()}")
        return jsonify({
            'success': False,
            'error': f'서버 오류가 발생했습니다: {str(e)}'
        }), 500


@blueprint.route('/api/months')
@login_required
def get_available_months():
    """데이터가 있는 모든 월 목록을 반환하는 API"""
    try:
        months = data_cache.get_all_months_with_data()
        return jsonify({
            'success': True,
            'months': months
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@blueprint.route('/<template>')
@login_required
def route_template(template):

    try:

        if not template.endswith('.html'):
            template += '.html'

        # Detect the current page
        segment = get_segment(request)

        # Serve the file (if exists) from app/templates/home/FILE.html
        return render_template("home/" + template, segment=segment)

    except TemplateNotFound:
        return render_template('home/page-404.html'), 404

    except:
        return render_template('home/page-500.html'), 500


# Helper - Extract current page name from request
def get_segment(request):

    try:

        segment = request.path.split('/')[-1]

        if segment == '':
            segment = 'index'

        return segment

    except:
        return None
