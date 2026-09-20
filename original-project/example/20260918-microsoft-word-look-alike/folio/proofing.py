from __future__ import annotations

import re
from dataclasses import dataclass

from spellchecker import SpellChecker

TOKEN_RE = re.compile(r"[A-Za-z]+(?:'[A-Za-z]+)?")
URL_RE = re.compile(r"(?:https?://|www\.)\S+", re.IGNORECASE)
FIELD_RE = re.compile(r"\{\{[^{}]*\}\}")
EXTRA_SPACES_RE = re.compile(r" {2,}")
WORD = "spelling"
GRAMMAR = "grammar"
STYLE = "style"

STYLE_RULES = (
    (re.compile(r"\bin order to\b", re.IGNORECASE), "to",
     "Style: 'in order to' can usually be shortened to 'to'."),
    (re.compile(r"\bvery unique\b", re.IGNORECASE), "unique",
     "Style: 'unique' is absolute; 'very unique' should be 'unique'."),
)

USER_DICTIONARY_KEY = "user_dictionary"


@dataclass
class ProofIssue:
    start: int
    length: int
    kind: str
    message: str
    word: str = ""
    replacement: str | None = None


def _utf16_offsets(text: str) -> list[int]:
    offsets = [0] * (len(text) + 1)
    count = 0
    for i, ch in enumerate(text):
        count += 2 if ord(ch) > 0xFFFF else 1
        offsets[i + 1] = count
    return offsets


class ProofingEngine:
    def __init__(self, store=None):
        self._spell = SpellChecker()
        self._memo: dict[str, bool] = {}
        self._user_words: set[str] = set()
        self._store = store
        if store is not None:
            saved = store.get_setting(USER_DICTIONARY_KEY, [])
            if isinstance(saved, list):
                self._user_words = {str(w).lower() for w in saved}

    def add_to_dictionary(self, word: str) -> None:
        word = word.strip().lower()
        if not word:
            return
        self._user_words.add(word)
        self._memo.pop(word, None)
        if self._store is not None:
            self._store.set_setting(USER_DICTIONARY_KEY, sorted(self._user_words))

    def _is_known(self, word: str) -> bool:
        lowered = word.lower()
        if lowered in self._user_words:
            return True
        cached = self._memo.get(lowered)
        if cached is not None:
            return cached
        known = not self._spell.unknown([lowered])
        self._memo[lowered] = known
        return known

    def suggestions(self, word: str, limit: int = 5) -> list[str]:
        candidates = self._spell.candidates(word.lower()) or set()
        candidates.discard(word.lower())
        ordered = sorted(
            candidates,
            key=lambda w: (
                -(self._spell.word_frequency[w] if w in self._spell.word_frequency else 0),
                w,
            ),
        )
        return ordered[:limit]

    def check(self, text: str) -> list[ProofIssue]:
        issues: list[ProofIssue] = []
        if not text:
            return issues
        offsets = _utf16_offsets(text)
        skip_spans: list[tuple[int, int]] = []
        for regex in (URL_RE, FIELD_RE):
            skip_spans.extend((m.start(), m.end()) for m in regex.finditer(text))
        skip_spans.sort()

        def skipped(start: int, end: int) -> bool:
            return any(s < end and start < e for s, e in skip_spans)

        tokens = list(TOKEN_RE.finditer(text))
        for match in tokens:
            word = match.group(0)
            if skipped(match.start(), match.end()):
                continue
            if len(word) > 1 and word.isupper():
                continue
            if word == "i":
                issues.append(ProofIssue(
                    start=offsets[match.start()],
                    length=offsets[match.end()] - offsets[match.start()],
                    kind=GRAMMAR,
                    message="Basic grammar: the pronoun 'I' should be capitalized.",
                    word=word,
                    replacement="I",
                ))
                continue
            if not self._is_known(word):
                issues.append(ProofIssue(
                    start=offsets[match.start()],
                    length=offsets[match.end()] - offsets[match.start()],
                    kind=WORD,
                    message=f"Possible spelling mistake: '{word}'.",
                    word=word,
                ))
        for first, second in zip(tokens, tokens[1:]):
            if skipped(first.start(), second.end()):
                continue
            between = text[first.end():second.start()]
            if "\n" in between or not between.strip() == "":
                continue
            if first.group(0).lower() == second.group(0).lower() and len(second.group(0)) > 1:
                issues.append(ProofIssue(
                    start=offsets[second.start()],
                    length=offsets[second.end()] - offsets[second.start()],
                    kind=GRAMMAR,
                    message=f"Basic grammar: repeated word '{second.group(0)}'.",
                    word=second.group(0),
                    replacement="",
                ))
        for match in EXTRA_SPACES_RE.finditer(text):
            if match.start() == 0 or text[match.start() - 1] == "\n":
                continue
            issues.append(ProofIssue(
                start=offsets[match.start()],
                length=offsets[match.end()] - offsets[match.start()],
                kind=GRAMMAR,
                message="Basic grammar: extra spaces.",
                word=match.group(0),
                replacement=" ",
            ))
        for regex, replacement, message in STYLE_RULES:
            for match in regex.finditer(text):
                if skipped(match.start(), match.end()):
                    continue
                issues.append(ProofIssue(
                    start=offsets[match.start()],
                    length=offsets[match.end()] - offsets[match.start()],
                    kind=STYLE,
                    message=message,
                    word=match.group(0),
                    replacement=replacement,
                ))
        issues.sort(key=lambda i: (i.start, i.length))
        return issues
