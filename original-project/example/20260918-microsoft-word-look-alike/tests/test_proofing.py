import pytest

from folio.proofing import GRAMMAR, STYLE, WORD, ProofingEngine


@pytest.fixture(scope="module")
def engine():
    return ProofingEngine()


def test_deliberate_typo(engine):
    issues = engine.check("This sentense has a mispeling.")
    words = {i.word for i in issues}
    assert "sentense" in words or "mispeling" in words


def test_repeated_word(engine):
    issues = engine.check("the the cat sat")
    assert any(i.kind == GRAMMAR and i.word == "the" for i in issues)


def test_lowercase_i(engine):
    issues = engine.check("i went home")
    assert any(i.kind == GRAMMAR and i.word == "i" and i.replacement == "I"
               for i in issues)


def test_corrected_text_removes_issues(engine):
    before = engine.check("This sentense is bad")
    after = engine.check("This sentence is good")
    assert len(after) < len(before)


def test_merge_placeholders_ignored(engine):
    issues = engine.check("Dear {{FirstName}}, hello {{Xyzq}}")
    assert not any("{{" in i.word for i in issues)


def test_urls_ignored(engine):
    issues = engine.check("visit https://www.examplllle.com/path today")
    assert not any("examplllle" in i.word for i in issues)


def test_utf16_offset_after_astral(engine):
    text = "\U0001D11E sentense"
    issues = engine.check(text)
    typo = [i for i in issues if i.word == "sentense"]
    assert typo
    expected = len("\U0001D11E ".encode("utf-16-le")) // 2
    assert typo[0].start == expected
    assert typo[0].length == len("sentense")


def test_suggestions(engine):
    suggestions = engine.suggestions("sentense", 5)
    assert isinstance(suggestions, list)
    assert "sentence" in suggestions


def test_add_to_dictionary(engine):
    word = "zxqwvbnm"
    assert any(i.word == word for i in engine.check(f"{word} ok"))
    engine.add_to_dictionary(word)
    assert not any(i.word == word for i in engine.check(f"{word} ok"))


def test_style_rules(engine):
    issues = engine.check("in order to win we need very unique ideas")
    kinds = {i.kind for i in issues}
    assert STYLE in kinds


def test_extra_spaces(engine):
    issues = engine.check("two  spaces here")
    assert any(i.word == "  " for i in issues)
