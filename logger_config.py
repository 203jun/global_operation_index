#!/usr/bin/env python3
"""
로깅 설정 모듈
다른 Python 파일들에서 재사용 가능한 로깅 설정을 제공합니다.
"""

import logging
import sys
import os
from datetime import datetime

def setup_logger(name=None, log_file=None, level=logging.INFO,
                 log_format=None, include_console=True):
    """
    로거를 설정하고 반환하는 함수

    Args:
        name (str, optional): 로거 이름. None이면 루트 로거 사용
        log_file (str, optional): 로그 파일 경로. None이면 파일 로깅 안함
        level (int): 로깅 레벨 (기본값: logging.INFO)
        log_format (str, optional): 로그 포맷. None이면 기본 포맷 사용
        include_console (bool): 콘솔 출력 포함 여부 (기본값: True)

    Returns:
        logging.Logger: 설정된 로거 객체
    """

    # 기본 로그 포맷 설정
    if log_format is None:
        log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'

    # 로거 생성
    logger = logging.getLogger(name)
    logger.setLevel(level)

    # 기존 핸들러 제거 (중복 방지)
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)

    # 포맷터 생성
    formatter = logging.Formatter(log_format)

    # 콘솔 핸들러 추가
    if include_console:
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(level)
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)

    # 파일 핸들러 추가
    if log_file:
        # 로그 디렉토리 생성
        log_dir = os.path.dirname(log_file)
        if log_dir and not os.path.exists(log_dir):
            os.makedirs(log_dir)

        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(level)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    return logger

def get_default_logger(module_name=None):
    """
    기본 설정으로 로거를 생성하는 함수

    Args:
        module_name (str, optional): 모듈 이름. None이면 __main__ 사용

    Returns:
        logging.Logger: 기본 설정된 로거 객체
    """
    if module_name is None:
        module_name = __name__

    # 기본 로그 파일명 생성
    timestamp = datetime.now().strftime("%Y%m%d")
    log_file = f"logs/{module_name}_{timestamp}.log"

    return setup_logger(
        name=module_name,
        log_file=log_file,
        level=logging.INFO,
        include_console=True
    )

def get_simple_logger(name=None):
    """
    간단한 로거를 생성하는 함수 (파일 로깅 없이 콘솔만)

    Args:
        name (str, optional): 로거 이름

    Returns:
        logging.Logger: 간단한 로거 객체
    """
    return setup_logger(
        name=name,
        log_file=None,
        level=logging.INFO,
        include_console=True
    )

# 사용 예시
if __name__ == "__main__":
    # 기본 로거 테스트
    logger = get_default_logger("test")
    logger.info("기본 로거 테스트")
    logger.warning("경고 메시지")
    logger.error("에러 메시지")

    # 간단한 로거 테스트
    simple_logger = get_simple_logger("simple_test")
    simple_logger.info("간단한 로거 테스트")
