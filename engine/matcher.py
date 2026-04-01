from dataclasses import dataclass
from enum import Enum
import re


class MatchMode(Enum):
    AND = "and"
    OR = "or"
    REGEX = "regex"


@dataclass
class MatchRule:
    """
    匹配规则数据类。

    Attributes:
        keywords:  关键词列表，AND/OR 模式使用。
        mode:      匹配模式（AND / OR / REGEX）。
        pattern:   正则表达式字符串，仅 REGEX 模式使用。
    """
    keywords: list
    mode: MatchMode
    pattern: str = None  # 仅 REGEX 模式使用

    # EXTENSION: NOT operator


def match_row(row_str: str, rule: MatchRule) -> bool:
    """
    将一行字符串与规则进行匹配，返回是否命中。

    AND:   所有 keywords 均出现在 row_str 中（大小写不敏感）。
    OR:    任一 keyword 出现。
    REGEX: re.search(pattern, row_str) 非空。

    # EXTENSION: NOT operator
    """
    if rule.mode == MatchMode.AND:
        return all(kw.lower() in row_str.lower() for kw in rule.keywords)
    elif rule.mode == MatchMode.OR:
        return any(kw.lower() in row_str.lower() for kw in rule.keywords)
    elif rule.mode == MatchMode.REGEX:
        if not rule.pattern:
            return False
        return re.search(rule.pattern, row_str) is not None
    return False
