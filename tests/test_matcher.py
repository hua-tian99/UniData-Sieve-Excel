"""tests/test_matcher.py — match_row 三模式 100% 覆盖率测试"""
import pytest
from engine.matcher import MatchRule, MatchMode, match_row


def test_and_mode_all_present():
    rule = MatchRule(keywords=["张三", "2026"], mode=MatchMode.AND)
    assert match_row("张三 2026/3/31 优秀", rule) is True


def test_and_mode_missing_one():
    rule = MatchRule(keywords=["张三", "缺勤"], mode=MatchMode.AND)
    assert match_row("张三 2026/3/31 优秀", rule) is False


def test_or_mode_one_present():
    rule = MatchRule(keywords=["缺勤", "优秀"], mode=MatchMode.OR)
    assert match_row("张三 2026/3/31 优秀", rule) is True


def test_or_mode_none_present():
    rule = MatchRule(keywords=["缺勤", "差评"], mode=MatchMode.OR)
    assert match_row("张三 2026/3/31 优秀", rule) is False


def test_regex_mode_match():
    rule = MatchRule(keywords=[], mode=MatchMode.REGEX, pattern=r"202\d/\d+/\d+")
    assert match_row("张三 2026/3/31", rule) is True


def test_regex_mode_no_match():
    rule = MatchRule(keywords=[], mode=MatchMode.REGEX, pattern=r"202\d/\d+/\d+")
    assert match_row("张三 李四 王五", rule) is False


def test_regex_mode_empty_pattern():
    rule = MatchRule(keywords=[], mode=MatchMode.REGEX, pattern=None)
    assert match_row("任意内容", rule) is False


def test_case_insensitive():
    rule = MatchRule(keywords=["ABC"], mode=MatchMode.AND)
    assert match_row("abc 123", rule) is True
