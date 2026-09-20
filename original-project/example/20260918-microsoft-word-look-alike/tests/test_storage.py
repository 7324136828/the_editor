import pytest

from folio.models import DocumentState
from folio.storage import LocalStore, StorageError


def test_recovery_round_trip(store):
    state = DocumentState(title="Recover me", html="<p>x</p>")
    store.save_recovery(state, None, dirty=True)
    entries = store.recoverable()
    assert len(entries) == 1
    assert entries[0]["doc_id"] == state.id
    loaded = store.load_recovery(state.id)
    assert loaded is not None
    recovered, path = loaded
    assert recovered.title == "Recover me"
    assert path is None


def test_mark_clean_removes_from_recoverable(store):
    state = DocumentState(title="t", html="<p>x</p>")
    store.save_recovery(state, None, dirty=True)
    store.mark_clean(state.id)
    assert store.recoverable() == []
    loaded = store.load_recovery(state.id)
    assert loaded is not None


def test_discard_recovery(store):
    state = DocumentState(title="t", html="<p>x</p>")
    store.save_recovery(state, None, dirty=True)
    store.discard_recovery(state.id)
    assert store.recoverable() == []
    assert store.load_recovery(state.id) is None


def test_clean_entry_not_recoverable(store):
    state = DocumentState(title="t", html="<p>x</p>")
    store.save_recovery(state, "some/path.docx", dirty=False)
    assert store.recoverable() == []


def test_versions(store):
    state = DocumentState(title="v", html="<p>one</p>")
    vid = store.add_version(state, "first")
    state.html = "<p>two</p>"
    store.add_version(state, "second")
    versions = store.versions(state.id)
    assert len(versions) == 2
    assert versions[0]["label"] == "second"
    restored = store.load_version(vid)
    assert "one" in restored.html
    with pytest.raises(StorageError):
        store.load_version(999999)


def test_settings_round_trip(store):
    assert store.get_setting("missing", "d") == "d"
    store.set_setting("author", "Zach")
    assert store.get_setting("author") == "Zach"
    store.set_setting("flag", True)
    assert store.get_setting("flag") is True


def test_reopen_store(store, tmp_path):
    state = DocumentState(title="persist", html="<p>x</p>")
    store.save_recovery(state, None, dirty=True)
    store.close()
    store2 = LocalStore(store.root)
    try:
        assert len(store2.recoverable()) == 1
    finally:
        store2.close()
