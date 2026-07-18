# src/utils/helpers.py
"""通用工具函数"""
from functools import lru_cache

from pypinyin import lazy_pinyin, pinyin, Style


def pinyin_key(name: str) -> str:
    """返回姓名的拼音字符串（小写），用于排序。"""
    if not name:
        return ""
    return "".join(lazy_pinyin(str(name))).lower()


@lru_cache(maxsize=4096)
def pinyin_readings(char: str) -> frozenset:
    """
    返回单个汉字的全部读音（无声调、小写），用于同音 / 模糊音比对。
    多音字（如“乐” le/yue、“行” xing/hang）返回所有候选读音。
    非汉字或取不到读音时回退为该字符本身。
    """
    if not char:
        return frozenset()
    readings = pinyin(char, heteronym=True, style=Style.NORMAL, errors="default")
    if not readings or not readings[0]:
        return frozenset({char.lower()})
    return frozenset(r.lower() for r in readings[0] if r)


def chars_homophone(a: str, b: str) -> bool:
    """两个汉字是否同音（多音字任一读音相交即算）。"""
    return bool(pinyin_readings(a) & pinyin_readings(b))


def _levenshtein(a, b) -> int:
    """标准 Levenshtein 编辑距离，适用于字符串或序列。"""
    if a == b:
        return 0
    if not a:
        return len(b)
    if not b:
        return len(a)
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, start=1):
        cur = [i]
        for j, cb in enumerate(b, start=1):
            cost = 0 if ca == cb else 1
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost))
        prev = cur
    return prev[-1]


def syllable_letter_distance(s1: str, s2: str) -> int:
    """两个拼音音节之间的字母级编辑距离。"""
    return _levenshtein(s1, s2)


def chars_sound_similar(a: str, b: str) -> bool:
    """
    两个汉字读音是否相近：同音，或任一读音对之间的字母编辑距离
    <= max(1, floor(L * 0.4))，其中 L 为该读音对中较长音节的字母数。
    用于模糊音判定（涵盖 zh/z、an/ang、in/ing、l/n 等单处混淆）。

    取 floor 而非 round：4 字母音节阈值为 1（拒绝 ming/peng 这类两处不同的异音），
    5~7 字母音节阈值为 2（容许 shuang/xiang/zhang 的双处混淆）。
    """
    readings_a = pinyin_readings(a)
    readings_b = pinyin_readings(b)
    if readings_a & readings_b:
        return True
    for ra in readings_a:
        for rb in readings_b:
            threshold = max(1, int(max(len(ra), len(rb)) * 0.4))
            if syllable_letter_distance(ra, rb) <= threshold:
                return True
    return False


def edit_distance(a, b) -> int:
    """姓名（字符序列）级编辑距离，供防打错字比对用。"""
    return _levenshtein(a, b)
