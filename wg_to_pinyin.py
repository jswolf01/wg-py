#!/usr/bin/env python3
"""
Wade-Giles to Pinyin Converter

This script reads a .docx file, detects Wade-Giles romanization,
and converts it to toneless Pinyin based on the standard conversion table.

Usage:
    python wg_to_pinyin.py input.docx output.docx
    python wg_to_pinyin.py input.docx  # outputs to input_pinyin.docx
"""

import re
import sys
import shutil
import zipfile
from pathlib import Path
from docx import Document


# Wade-Giles to Pinyin conversion dictionary
# Based on "PINYIN CONVERSION SPECIFICATIONS - DICTIONARY STD: Standard pinyin conversion table"
#
# Notes on Wade-Giles orthography:
# - Apostrophe (') indicates aspiration: ch' -> ch, k' -> k, p' -> p, t' -> t, ts' -> c
# - Without apostrophe: ch -> zh, k -> g, p -> b, t -> d, ts -> z
# - ü is used after certain consonants (ch', hs, l, n, y)
# - Some syllables have variant spellings with diacritics (circumflex, breve) in original sources

# Core conversion dictionary (lowercase)
# Format: wade_giles -> pinyin
WG_TO_PINYIN = {
    # A
    'a': 'a',
    'ai': 'ai',
    'an': 'an',
    'ang': 'ang',
    'ao': 'ao',

    # CH (unaspirated -> zh)
    'cha': 'zha',
    'chai': 'zhai',
    'chan': 'zhan',
    'chang': 'zhang',
    'chao': 'zhao',
    'che': 'zhe',
    'chen': 'zhen',
    'cheng': 'zheng',
    'chi': 'ji',
    'chia': 'jia',
    'chiang': 'jiang',
    'chiao': 'jiao',
    'chieh': 'jie',
    'chien': 'jian',
    'chih': 'zhi',
    'chin': 'jin',
    'ching': 'jing',
    'chiu': 'jiu',
    'chiung': 'jiong',
    'cho': 'zhuo',
    'chou': 'zhou',
    'chu': 'zhu',
    'chua': 'zhua',
    'chuai': 'zhuai',
    'chuan': 'zhuan',
    'chuang': 'zhuang',
    'chui': 'zhui',
    'chun': 'zhun',
    'chung': 'zhong',
    'chü': 'ju',
    'chüan': 'juan',
    'chüeh': 'jue',
    'chün': 'jun',

    # CH' (aspirated -> ch/q)
    "ch'a": 'cha',
    "ch'ai": 'chai',
    "ch'an": 'chan',
    "ch'ang": 'chang',
    "ch'ao": 'chao',
    "ch'e": 'che',
    "ch'en": 'chen',
    "ch'eng": 'cheng',
    "ch'i": 'qi',
    "ch'ia": 'qia',
    "ch'iang": 'qiang',
    "ch'iao": 'qiao',
    "ch'ieh": 'qie',
    "ch'ien": 'qian',
    "ch'ih": 'chi',
    "ch'in": 'qin',
    "ch'ing": 'qing',
    "ch'iu": 'qiu',
    "ch'iung": 'qiong',
    "ch'o": 'chuo',
    "ch'ou": 'chou',
    "ch'u": 'chu',
    "ch'uai": 'chuai',
    "ch'uan": 'chuan',
    "ch'uang": 'chuang',
    "ch'ui": 'chui',
    "ch'un": 'chun',
    "ch'ung": 'chong',
    "ch'ü": 'qu',
    "ch'üan": 'quan',
    "ch'üeh": 'que',
    "ch'ün": 'qun',

    # E
    'en': 'en',
    'erh': 'er',
    'er': 'er',  # variant

    # F
    'fa': 'fa',
    'fan': 'fan',
    'fang': 'fang',
    'fei': 'fei',
    'fen': 'fen',
    'feng': 'feng',
    'fo': 'fo',
    'fou': 'fou',
    'fu': 'fu',

    # H
    'ha': 'ha',
    'hai': 'hai',
    'han': 'han',
    'hang': 'hang',
    'hao': 'hao',
    'hei': 'hei',
    'hen': 'hen',
    'heng': 'heng',
    'ho': 'he',
    'hou': 'hou',
    'hu': 'hu',
    'hua': 'hua',
    'huai': 'huai',
    'huan': 'huan',
    'huang': 'huang',
    'hui': 'hui',
    'hun': 'hun',
    'hung': 'hong',
    'huo': 'huo',

    # HS -> x
    'hsi': 'xi',
    'hsia': 'xia',
    'hsiang': 'xiang',
    'hsiao': 'xiao',
    'hsieh': 'xie',
    'hsien': 'xian',
    'hsin': 'xin',
    'hsing': 'xing',
    'hsiu': 'xiu',
    'hsiung': 'xiong',
    'hsü': 'xu',
    'hsüan': 'xuan',
    'hsüeh': 'xue',
    'hsün': 'xun',

    # I
    'i': 'yi',

    # J -> r
    'jan': 'ran',
    'jang': 'rang',
    'jao': 'rao',
    'je': 're',
    'jen': 'ren',
    'jeng': 'reng',
    'jih': 'ri',
    'jo': 'ruo',
    'jou': 'rou',
    'ju': 'ru',
    'juan': 'ruan',
    'jui': 'rui',
    'jun': 'run',
    'jung': 'rong',

    # K (unaspirated -> g)
    'ka': 'ga',
    'kai': 'gai',
    'kan': 'gan',
    'kang': 'gang',
    'kao': 'gao',
    'kei': 'gei',
    'ken': 'gen',
    'keng': 'geng',
    'ko': 'ge',
    'kou': 'gou',
    'ku': 'gu',
    'kua': 'gua',
    'kuai': 'guai',
    'kuan': 'guan',
    'kuang': 'guang',
    'kuei': 'gui',
    'kun': 'gun',
    'kung': 'gong',
    'kuo': 'guo',

    # K' (aspirated -> k)
    "k'a": 'ka',
    "k'ai": 'kai',
    "k'an": 'kan',
    "k'ang": 'kang',
    "k'ao": 'kao',
    "k'en": 'ken',
    "k'eng": 'keng',
    "k'o": 'ke',
    "k'ou": 'kou',
    "k'u": 'ku',
    "k'ua": 'kua',
    "k'uai": 'kuai',
    "k'uan": 'kuan',
    "k'uang": 'kuang',
    "k'uei": 'kui',
    "k'un": 'kun',
    "k'ung": 'kong',
    "k'uo": 'kuo',

    # L
    'la': 'la',
    'lai': 'lai',
    'lan': 'lan',
    'lang': 'lang',
    'lao': 'lao',
    'le': 'le',
    'lei': 'lei',
    'leng': 'leng',
    'li': 'li',
    'liang': 'liang',
    'liao': 'liao',
    'lieh': 'lie',
    'lien': 'lian',
    'lin': 'lin',
    'ling': 'ling',
    'liu': 'liu',
    'lo': 'luo',
    'lou': 'lou',
    'lu': 'lu',
    'luan': 'luan',
    'lun': 'lun',
    'lung': 'long',
    'lü': 'lü',
    'lüan': 'luan',
    'lüeh': 'lue',

    # M
    'ma': 'ma',
    'mai': 'mai',
    'man': 'man',
    'mang': 'mang',
    'mao': 'mao',
    'mei': 'mei',
    'men': 'men',
    'meng': 'meng',
    'mi': 'mi',
    'miao': 'miao',
    'mieh': 'mie',
    'mien': 'mian',
    'min': 'min',
    'ming': 'ming',
    'miu': 'miu',
    'mo': 'mo',
    'mou': 'mou',
    'mu': 'mu',

    # N
    'na': 'na',
    'nai': 'nai',
    'nan': 'nan',
    'nang': 'nang',
    'nao': 'nao',
    'nei': 'nei',
    'nen': 'nen',
    'neng': 'neng',
    'ni': 'ni',
    'niang': 'niang',
    'niao': 'niao',
    'nieh': 'nie',
    'nien': 'nian',
    'nin': 'nin',
    'ning': 'ning',
    'niu': 'niu',
    'no': 'nuo',
    'nu': 'nu',
    'nuan': 'nuan',
    'nung': 'nong',
    'nü': 'nü',
    'nüeh': 'nue',

    # O
    'o': 'e',
    'ou': 'ou',

    # P (unaspirated -> b)
    'pa': 'ba',
    'pai': 'bai',
    'pan': 'ban',
    'pang': 'bang',
    'pao': 'bao',
    'pei': 'bei',
    'pen': 'ben',
    'peng': 'beng',
    'pi': 'bi',
    'piao': 'biao',
    'pieh': 'bie',
    'pien': 'bian',
    'pin': 'bin',
    'ping': 'bing',
    'po': 'bo',
    'pu': 'bu',

    # P' (aspirated -> p)
    "p'a": 'pa',
    "p'ai": 'pai',
    "p'an": 'pan',
    "p'ang": 'pang',
    "p'ao": 'pao',
    "p'ei": 'pei',
    "p'en": 'pen',
    "p'eng": 'peng',
    "p'i": 'pi',
    "p'iao": 'piao',
    "p'ieh": 'pie',
    "p'ien": 'pian',
    "p'in": 'pin',
    "p'ing": 'ping',
    "p'o": 'po',
    "p'ou": 'pou',
    "p'u": 'pu',

    # S
    'sa': 'sa',
    'sai': 'sai',
    'san': 'san',
    'sang': 'sang',
    'sao': 'sao',
    'se': 'se',
    'sen': 'sen',
    'seng': 'seng',
    'sha': 'sha',
    'shai': 'shai',
    'shan': 'shan',
    'shang': 'shang',
    'shao': 'shao',
    'she': 'she',
    'shen': 'shen',
    'sheng': 'sheng',
    'shih': 'shi',
    'shou': 'shou',
    'shu': 'shu',
    'shua': 'shua',
    'shuai': 'shuai',
    'shuan': 'shuan',
    'shuang': 'shuang',
    'shui': 'shui',
    'shun': 'shun',
    'shuo': 'shuo',
    'so': 'suo',
    'sou': 'sou',
    'ssu': 'si',
    'su': 'su',
    'suan': 'suan',
    'sui': 'sui',
    'sun': 'sun',
    'sung': 'song',
    'szu': 'si',

    # T (unaspirated -> d)
    'ta': 'da',
    'tai': 'dai',
    'tan': 'dan',
    'tang': 'dang',
    'tao': 'dao',
    'te': 'de',
    'teng': 'deng',
    'ti': 'di',
    'tiao': 'diao',
    'tieh': 'die',
    'tien': 'dian',
    'ting': 'ding',
    'tiu': 'diu',
    'to': 'duo',
    'tou': 'dou',
    'tu': 'du',
    'tuan': 'duan',
    'tui': 'dui',
    'tun': 'dun',
    'tung': 'dong',

    # T' (aspirated -> t)
    "t'a": 'ta',
    "t'ai": 'tai',
    "t'an": 'tan',
    "t'ang": 'tang',
    "t'ao": 'tao',
    "t'e": 'te',
    "t'eng": 'teng',
    "t'i": 'ti',
    "t'iao": 'tiao',
    "t'ieh": 'tie',
    "t'ien": 'tian',
    "t'ing": 'ting',
    "t'o": 'tuo',
    "t'ou": 'tou',
    "t'u": 'tu',
    "t'uan": 'tuan',
    "t'ui": 'tui',
    "t'un": 'tun',
    "t'ung": 'tong',

    # TS (unaspirated -> z)
    'tsa': 'za',
    'tsai': 'zai',
    'tsan': 'zan',
    'tsang': 'zang',
    'tsao': 'zao',
    'tse': 'ze',
    'tsei': 'zei',
    'tsen': 'zen',
    'tseng': 'zeng',
    'tso': 'zuo',
    'tsou': 'zou',
    'tsu': 'zu',
    'tsuan': 'zuan',
    'tsui': 'zui',
    'tsun': 'zun',
    'tsung': 'zong',
    'tzu': 'zi',

    # TS' (aspirated -> c)
    "ts'a": 'ca',
    "ts'ai": 'cai',
    "ts'an": 'can',
    "ts'ang": 'cang',
    "ts'ao": 'cao',
    "ts'e": 'ce',
    "ts'en": 'cen',
    "ts'eng": 'ceng',
    "ts'o": 'cuo',
    "ts'ou": 'cou',
    "ts'u": 'cu',
    "ts'uan": 'cuan',
    "ts'ui": 'cui',
    "ts'un": 'cun',
    "ts'ung": 'cong',
    "tz'u": 'ci',

    # W
    'wa': 'wa',
    'wai': 'wai',
    'wan': 'wan',
    'wang': 'wang',
    'wei': 'wei',
    'wen': 'wen',
    'weng': 'weng',
    'wo': 'wo',
    'wu': 'wu',

    # Y
    'ya': 'ya',
    'yai': 'yai',
    'yang': 'yang',
    'yao': 'yao',
    'yeh': 'ye',
    'yen': 'yan',
    'yin': 'yin',
    'ying': 'ying',
    'yo': 'yo',
    'yu': 'you',
    'yung': 'yong',
    'yü': 'yu',
    'yüan': 'yuan',
    'yüeh': 'yue',
    'yün': 'yun',
}

# Build additional variants for ü spelled as 'u' after certain consonants
# In some sources, ü is written as plain 'u' after j, q, x, y (in WG: ch', hs, y)
# We already handle chü, hsü, yü forms above
# Also add variants without the umlaut for broader matching

UMLAUT_VARIANTS = {
    # chü variants
    'chu': 'ju',  # This conflicts with chu -> zhu, so we handle context
    # We keep chü -> ju as the primary mapping

    # hsü variants (when written as hsu)
    'hsu': 'xu',
    'hsuan': 'xuan',
    'hsueh': 'xue',
    'hsun': 'xun',

    # yü variants (when written as yu without umlaut)
    # yu -> you is already in main dict
    'yuan': 'yuan',
    'yueh': 'yue',
    'yun': 'yun',

    # lü variants
    'lu': 'lu',  # Standard mapping
    'lueh': 'lue',

    # nü variants
    'nueh': 'nue',
}

# Merge umlaut variants (only add if not conflicting with existing)
for wg, py in UMLAUT_VARIANTS.items():
    if wg not in WG_TO_PINYIN:
        WG_TO_PINYIN[wg] = py


# Postal romanizations and other common variant spellings
# These are not standard Wade-Giles but appear frequently in older texts
POSTAL_ROMANIZATIONS = {
    # Province names
    'kwangsi': 'guangxi',
    'kwangtung': 'guangdong',
    'fukien': 'fujian',
    'chekiang': 'zhejiang',
    'kiangsi': 'jiangxi',
    'kiangsu': 'jiangsu',
    'shansi': 'shanxi',
    'shensi': 'shaanxi',
    'szechwan': 'sichuan',
    'szechuan': 'sichuan',
    'hopei': 'hebei',
    'hopeh': 'hebei',
    'honan': 'henan',
    'hupei': 'hubei',
    'hupeh': 'hubei',
    'hunan': 'hunan',
    'kansu': 'gansu',
    'kweichow': 'guizhou',
    'yunnan': 'yunnan',
    'anhwei': 'anhui',
    'chihli': 'zhili',
    'fengtien': 'fengtian',
    'manchuria': 'manchuria',  # Keep as is (not Chinese)

    # Major cities
    'peking': 'beijing',
    'peiping': 'beiping',
    'nanking': 'nanjing',
    'canton': 'guangzhou',
    'tientsin': 'tianjin',
    'tsingtao': 'qingdao',
    'chungking': 'chongqing',
    'sian': 'xian',
    'sinkiang': 'xinjiang',
    'tsinghai': 'qinghai',
    'ningsia': 'ningxia',
    'suiyuan': 'suiyuan',

    # Rivers and geographic features
    'yangtze': 'yangzi',
    'yangtse': 'yangzi',

    # Common "kw" forms (Cantonese influence)
    'kwang': 'guang',
    'kwan': 'guan',
    'kwai': 'guai',
    'kwei': 'gui',
}

# Add postal romanizations to main dictionary
for postal, pinyin in POSTAL_ROMANIZATIONS.items():
    if postal not in WG_TO_PINYIN:
        WG_TO_PINYIN[postal] = pinyin


# PDF artifact: "ii" often represents ü (u with umlaut) in converted documents
# This mapping handles syllables where ü appears as "ii"
II_TO_UMLAUT_MAPPINGS = {
    # ch + ii (aspirated, with umlaut) -> q + u
    "ch'ii": 'qu',
    "ch'iian": 'quan',
    "ch'iieh": 'que',
    "ch'iin": 'qun',

    # ch + ii (unaspirated, with umlaut) -> j + u
    'chii': 'ju',
    'chiian': 'juan',
    'chiieh': 'jue',
    'chiin': 'jun',

    # hs + ii -> x + u
    'hsii': 'xu',
    'hsiian': 'xuan',
    'hsiieh': 'xue',
    'hsiin': 'xun',

    # l + ii -> l + ü
    'lii': 'lü',
    'liian': 'luan',
    'liieh': 'lue',

    # n + ii -> n + ü
    'nii': 'nü',
    'niieh': 'nue',

    # y + ii -> y + u
    'yii': 'yu',
    'yiian': 'yuan',
    'yiieh': 'yue',
    'yiin': 'yun',
}

# Add ii mappings to main dictionary
for ii_form, pinyin in II_TO_UMLAUT_MAPPINGS.items():
    if ii_form not in WG_TO_PINYIN:
        WG_TO_PINYIN[ii_form] = pinyin


# Common English words that happen to match Wade-Giles patterns
# These should NOT be converted unless in a clearly Chinese context
# (e.g., hyphenated with other syllables or part of a proper name sequence)
# Only include words where:
# 1. The WG -> PY conversion changes the spelling, AND
# 2. The word is very common in English
ENGLISH_EXCLUSIONS_LOWERCASE = {
    # Very common English function words (lowercase only)
    'to',      # to -> duo (very common English word)
    'no',      # no -> nuo (very common English word)
    'so',      # so -> suo (very common English word)

    # English words where conversion would cause confusion (lowercase only)
    # Note: Capitalized versions may be Chinese proper names, so we only
    # exclude lowercase forms
    'hung',    # hung -> hong (past tense of "hang")
    'sung',    # sung -> song (past tense of "sing")
    'lung',    # lung -> long (body organ)
    'tang',    # tang -> dang (sharp taste/flavor)
    'tan',     # tan -> dan (skin color/tanning)
    'pan',     # pan -> ban (cooking vessel)
    'pen',     # pen -> ben (writing instrument)
    'pin',     # pin -> bin (fastening device)
    'ping',    # ping -> bing (network ping, sound)
    'ting',    # ting -> ding (sound)
}

# Words that should be excluded even when capitalized
# (these are rarely Chinese names)
ENGLISH_EXCLUSIONS_ANY_CASE = {
    'to',
    'no',
    'so',
}

# Syllables that should ONLY be converted in hyphenated Chinese contexts
CONTEXT_SENSITIVE = {
    'i',       # i -> yi (but also English "I")
    'a',       # a -> a (English article)
    'o',       # o -> e (English exclamation)
}


def apply_case(source: str, target: str) -> str:
    """Apply the capitalization pattern from source to target."""
    if not source or not target:
        return target

    result = []
    target_idx = 0

    for i, char in enumerate(source):
        if target_idx >= len(target):
            break

        # Skip apostrophes and hyphens in source when matching
        if char in "''-":
            continue

        if char.isupper():
            result.append(target[target_idx].upper())
        else:
            result.append(target[target_idx].lower())
        target_idx += 1

    # Append any remaining characters from target
    if target_idx < len(target):
        # Determine case from last character of source
        if source and source[-1].isupper():
            result.append(target[target_idx:].upper())
        else:
            result.append(target[target_idx:].lower())

    return ''.join(result)


def normalize_apostrophe(text: str) -> str:
    """Normalize various apostrophe characters to standard apostrophe."""
    # Various apostrophe-like characters
    apostrophes = [''', ''', '`', '´', 'ʼ', 'ʻ', '\u02bc', '\u02bb']
    for apos in apostrophes:
        text = text.replace(apos, "'")
    return text


def build_regex_pattern():
    """Build a regex pattern to match all Wade-Giles syllables."""
    # Sort by length (longest first) to ensure longer matches take precedence
    all_syllables = sorted(WG_TO_PINYIN.keys(), key=len, reverse=True)

    # Escape special regex characters and handle apostrophes
    escaped = []
    for syl in all_syllables:
        # Replace apostrophe with pattern matching multiple apostrophe types
        if "'" in syl:
            parts = syl.split("'")
            pattern = "[''`´ʼʻ]?".join(re.escape(p) for p in parts)
            escaped.append(pattern)
        else:
            escaped.append(re.escape(syl))

    # Build pattern with word boundaries
    # Match syllables that are standalone words or connected with hyphens
    pattern = r'\b(' + '|'.join(escaped) + r')\b'

    return pattern


def convert_syllable(match, case_sensitive=True):
    """Convert a matched Wade-Giles syllable to Pinyin."""
    original = match.group(0)
    normalized = normalize_apostrophe(original.lower())

    # Handle ü character (might appear as ü or u with combining umlaut)
    normalized = normalized.replace('ü', 'ü')  # Normalize composed form

    if normalized in WG_TO_PINYIN:
        pinyin = WG_TO_PINYIN[normalized]
        if case_sensitive:
            return apply_case(original, pinyin)
        return pinyin

    return original


class WadeGilesToPinyinConverter:
    """Converter class for Wade-Giles to Pinyin conversion."""

    def __init__(self):
        self.pattern = None
        self._compile_pattern()

    def _compile_pattern(self):
        """Compile the regex pattern for syllable matching."""
        all_syllables = sorted(WG_TO_PINYIN.keys(), key=len, reverse=True)

        patterns = []
        for syl in all_syllables:
            if "'" in syl:
                # Handle apostrophe variants
                parts = syl.split("'")
                pattern = "[''`´ʼʻ']".join(re.escape(p) for p in parts)
                patterns.append(pattern)
            elif 'ü' in syl:
                # Handle umlaut variants
                pattern = re.escape(syl).replace(re.escape('ü'), '[üu]')
                patterns.append(pattern)
            else:
                patterns.append(re.escape(syl))

        # Word boundary pattern
        self.pattern = re.compile(
            r'(?<![a-zA-Z])(' + '|'.join(patterns) + r')(?![a-zA-Z])',
            re.IGNORECASE
        )

    def _convert_single_syllable(self, original: str, aggressive: bool = False,
                                   in_hyphenated_sequence: bool = False) -> str:
        """Convert a single Wade-Giles syllable to Pinyin."""
        normalized = normalize_apostrophe(original.lower())
        normalized = normalized.replace('ü', 'ü')

        # Skip common English words unless in aggressive mode or in a hyphenated sequence
        if not aggressive and not in_hyphenated_sequence:
            if normalized in ENGLISH_EXCLUSIONS_ANY_CASE:
                return original
            if normalized in ENGLISH_EXCLUSIONS_LOWERCASE and original.islower():
                return original

        # Skip context-sensitive syllables unless in hyphenated sequence
        if not aggressive and not in_hyphenated_sequence and normalized in CONTEXT_SENSITIVE:
            return original

        if normalized in WG_TO_PINYIN:
            pinyin = WG_TO_PINYIN[normalized]
            return apply_case(original, pinyin)
        return original

    def _convert_hyphenated_sequence(self, sequence: str, aggressive: bool = False) -> str:
        """
        Convert a hyphenated sequence of Wade-Giles syllables to Pinyin.
        Removes hyphens in the output (e.g., "Tse-tung" -> "Zedong").
        """
        parts = sequence.split('-')
        converted_parts = []
        all_are_wg_syllables = True
        any_converted = False

        for part in parts:
            if not part:
                continue
            # Check if this part matches a WG syllable
            normalized = normalize_apostrophe(part.lower()).replace('ü', 'ü')
            if normalized in WG_TO_PINYIN:
                converted = self._convert_single_syllable(part, aggressive=True,
                                                          in_hyphenated_sequence=True)
                converted_parts.append(converted)
                if converted.lower() != part.lower():
                    any_converted = True
            else:
                # Not a WG syllable, keep original
                converted_parts.append(part)
                all_are_wg_syllables = False

        # Remove hyphens if:
        # 1. At least one syllable was actually converted (spelling changed), OR
        # 2. ALL parts are valid WG syllables (likely a Chinese name even if spellings match)
        if any_converted or all_are_wg_syllables:
            return ''.join(converted_parts)
        else:
            return sequence

    def convert_text(self, text: str, aggressive: bool = False) -> str:
        """
        Convert Wade-Giles text to Pinyin.

        Args:
            text: Input text containing Wade-Giles romanization
            aggressive: If True, convert all matches including common English words.
                       If False (default), skip common English words to reduce false positives.

        Returns:
            Text with Wade-Giles converted to Pinyin
        """
        if not text:
            return text

        # First, handle hyphenated sequences (e.g., "Tse-tung" -> "Zedong")
        # Match sequences of syllables connected by hyphens
        hyphen_pattern = re.compile(
            r'\b([a-zA-ZüÜ\'\'`´ʼʻ]+(?:-[a-zA-ZüÜ\'\'`´ʼʻ]+)+)\b'
        )

        def hyphen_replacer(match):
            return self._convert_hyphenated_sequence(match.group(1), aggressive)

        text = hyphen_pattern.sub(hyphen_replacer, text)

        # Then handle standalone syllables
        def replacer(match):
            original = match.group(1)
            return self._convert_single_syllable(original, aggressive, in_hyphenated_sequence=False)

        return self.pattern.sub(replacer, text)

    def _process_textboxes_in_xml(self, xml_content: bytes, aggressive: bool = False) -> bytes:
        """
        Process text boxes in the document XML using efficient regex-based replacement.

        This approach is much faster and more memory-efficient than DOM parsing
        for large documents. It finds all <w:t>...</w:t> elements within
        <w:txbxContent> sections and applies conversion to their text content.

        Args:
            xml_content: The raw XML content of document.xml
            aggressive: If True, convert all matches including common English words

        Returns:
            Modified XML content as bytes
        """
        # Decode to string for processing
        xml_str = xml_content.decode('utf-8')

        conversion_count = 0

        # Find all txbxContent sections and process text within them
        # Pattern to find text box content sections
        txbx_pattern = re.compile(
            r'(<w:txbxContent[^>]*>)(.*?)(</w:txbxContent>)',
            re.DOTALL
        )

        def process_textbox_content(match):
            nonlocal conversion_count
            start_tag = match.group(1)
            content = match.group(2)
            end_tag = match.group(3)

            # Within textbox content, find and convert text in <w:t> elements
            # Pattern for <w:t>text</w:t> or <w:t xml:space="preserve">text</w:t>
            text_pattern = re.compile(r'(<w:t[^>]*>)([^<]*)(</w:t>)')

            def convert_text_element(text_match):
                nonlocal conversion_count
                open_tag = text_match.group(1)
                text = text_match.group(2)
                close_tag = text_match.group(3)

                if text:
                    converted = self.convert_text(text, aggressive=aggressive)
                    if converted != text:
                        conversion_count += 1
                        return open_tag + converted + close_tag
                return text_match.group(0)

            processed_content = text_pattern.sub(convert_text_element, content)
            return start_tag + processed_content + end_tag

        modified_xml = txbx_pattern.sub(process_textbox_content, xml_str)

        if conversion_count > 0:
            print(f"  Converted {conversion_count} text elements in text boxes")

        return modified_xml.encode('utf-8')

    def _process_all_text_in_xml(self, xml_content: bytes, aggressive: bool = False) -> bytes:
        """
        Process ALL text in the document XML using regex-based replacement.

        This method processes all <w:t> elements in the document, not just those
        in text boxes. Used as a fallback when python-docx fails to load large documents.

        Args:
            xml_content: The raw XML content of document.xml
            aggressive: If True, convert all matches including common English words

        Returns:
            Modified XML content as bytes
        """
        # Decode to string for processing
        xml_str = xml_content.decode('utf-8')

        conversion_count = 0

        # Pattern for all <w:t>text</w:t> elements
        text_pattern = re.compile(r'(<w:t[^>]*>)([^<]*)(</w:t>)')

        def convert_text_element(text_match):
            nonlocal conversion_count
            open_tag = text_match.group(1)
            text = text_match.group(2)
            close_tag = text_match.group(3)

            if text:
                converted = self.convert_text(text, aggressive=aggressive)
                if converted != text:
                    conversion_count += 1
                    return open_tag + converted + close_tag
            return text_match.group(0)

        modified_xml = text_pattern.sub(convert_text_element, xml_str)

        print(f"  Converted {conversion_count} text elements")

        return modified_xml.encode('utf-8')

    def _convert_docx_via_xml(self, input_path: Path, output_path: Path,
                               aggressive: bool = False) -> str:
        """
        Convert a .docx file using direct XML manipulation only.

        This is used as a fallback when python-docx fails to load the document
        (e.g., due to XML parser limits with very large documents).

        Args:
            input_path: Path to input .docx file
            output_path: Path for output file
            aggressive: If True, convert all matches including common English words

        Returns:
            Path to the output file
        """
        # Copy the input file to output first
        shutil.copy2(input_path, output_path)

        # Process the XML directly
        with zipfile.ZipFile(output_path, 'r') as zf:
            xml_content = zf.read('word/document.xml')

            # Process ALL text in the document
            modified_xml = self._process_all_text_in_xml(xml_content, aggressive)

            # Read all other files
            other_files = {}
            for item in zf.namelist():
                if item != 'word/document.xml':
                    other_files[item] = zf.read(item)

        # Write the modified docx
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.writestr('word/document.xml', modified_xml)
            for name, content in other_files.items():
                zf.writestr(name, content)

        return str(output_path)

    def convert_docx(self, input_path: str, output_path: str = None,
                      aggressive: bool = False) -> str:
        """
        Convert Wade-Giles in a .docx file to Pinyin.

        This method processes:
        - Regular paragraphs
        - Tables
        - Headers and footers
        - Text boxes (via direct XML manipulation)

        For very large documents that exceed XML parser limits, it automatically
        falls back to direct XML processing.

        Args:
            input_path: Path to input .docx file
            output_path: Path for output file (optional)
            aggressive: If True, convert all matches including common English words

        Returns:
            Path to the output file
        """
        input_path = Path(input_path)

        if output_path is None:
            output_path = input_path.parent / f"{input_path.stem}_pinyin{input_path.suffix}"
        else:
            output_path = Path(output_path)

        # Try to process with python-docx first
        try:
            print("Processing regular content (paragraphs, tables, headers/footers)...")
            doc = Document(str(input_path))

            # Process paragraphs
            for para in doc.paragraphs:
                for run in para.runs:
                    if run.text:
                        run.text = self.convert_text(run.text, aggressive=aggressive)

            # Process tables
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            for run in para.runs:
                                if run.text:
                                    run.text = self.convert_text(run.text, aggressive=aggressive)

            # Process headers and footers
            for section in doc.sections:
                header = section.header
                for para in header.paragraphs:
                    for run in para.runs:
                        if run.text:
                            run.text = self.convert_text(run.text, aggressive=aggressive)

                footer = section.footer
                for para in footer.paragraphs:
                    for run in para.runs:
                        if run.text:
                            run.text = self.convert_text(run.text, aggressive=aggressive)

            # Save intermediate result
            doc.save(str(output_path))

            # Now process text boxes by directly modifying the XML
            print("Processing text boxes...")
            self._process_textboxes_in_docx(output_path, aggressive)

        except Exception as e:
            # Check if this is an XML parser error (common with large PDF-converted docs)
            error_msg = str(e)
            if 'XMLSyntaxError' in type(e).__name__ or 'AttValue' in error_msg or 'lxml' in error_msg:
                print(f"Document too large for standard processing. Using direct XML mode...")
                print("Processing all text via XML manipulation...")
                return self._convert_docx_via_xml(input_path, output_path, aggressive)
            else:
                # Re-raise other errors
                raise

        return str(output_path)

    def _process_textboxes_in_docx(self, docx_path: Path, aggressive: bool = False):
        """
        Process text boxes in a .docx file by modifying its XML directly.

        Args:
            docx_path: Path to the .docx file to modify in place
            aggressive: If True, convert all matches including common English words
        """
        docx_path = Path(docx_path)

        # Read the docx as a zip file
        with zipfile.ZipFile(docx_path, 'r') as zf:
            # Read document.xml
            xml_content = zf.read('word/document.xml')

            # Check if there are any text boxes
            if b'txbxContent' not in xml_content:
                print("  No text boxes found in document")
                return

            # Process the XML
            modified_xml = self._process_textboxes_in_xml(xml_content, aggressive)

            # Read all other files
            other_files = {}
            for item in zf.namelist():
                if item != 'word/document.xml':
                    other_files[item] = zf.read(item)

        # Write the modified docx
        with zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            # Write modified document.xml
            zf.writestr('word/document.xml', modified_xml)

            # Write all other files unchanged
            for name, content in other_files.items():
                zf.writestr(name, content)


def main():
    """Main entry point for command-line usage."""
    import argparse

    parser = argparse.ArgumentParser(
        description='Convert Wade-Giles romanization to Pinyin in .docx files.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python wg_to_pinyin.py document.docx
  python wg_to_pinyin.py input.docx output.docx
  python wg_to_pinyin.py input.docx -o output.docx --aggressive

Notes:
  - By default, common English words (to, no, so, lung, tang, etc.) are NOT
    converted to avoid false positives.
  - Capitalized versions of ambiguous words ARE converted, as they may be
    Chinese proper names (e.g., "Sung Dynasty" -> "Song Dynasty").
  - Use --aggressive to convert ALL matches including common English words.
        """
    )

    parser.add_argument('input', help='Input .docx file')
    parser.add_argument('output', nargs='?', help='Output .docx file (default: input_pinyin.docx)')
    parser.add_argument('-o', '--output-file', dest='output_file',
                        help='Output .docx file (alternative to positional argument)')
    parser.add_argument('-a', '--aggressive', action='store_true',
                        help='Convert all matches, including common English words')

    args = parser.parse_args()

    input_file = args.input
    output_file = args.output or args.output_file

    if not Path(input_file).exists():
        print(f"Error: Input file '{input_file}' not found.")
        sys.exit(1)

    converter = WadeGilesToPinyinConverter()

    if args.aggressive:
        # For aggressive mode, we need to modify convert_docx or pass a flag
        # Let's modify the converter class to support this
        output_path = converter.convert_docx(input_file, output_file, aggressive=True)
    else:
        output_path = converter.convert_docx(input_file, output_file)

    print(f"Conversion complete. Output saved to: {output_path}")


if __name__ == '__main__':
    main()
