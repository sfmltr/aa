"""Mail search (with a fake Outlook), cache, model matching and path helpers."""
import datetime as dt
import threading


from ncr_panel import outlook_mail as om
from ncr_panel.catia import match_models
from ncr_panel.config import OutlookSettings
from ncr_panel.utils import norm_path


# ---- a tiny fake of the Outlook object model used by outlook_mail.search -----------
class _Row:
    def __init__(self, vals):
        self._v = vals

    def GetValues(self):
        return self._v


class _Columns:
    def RemoveAll(self):
        pass

    def Add(self, _name):
        pass


class _Table:
    def __init__(self, rows):
        self._rows = list(rows)
        self.Columns = _Columns()

    @property
    def EndOfTable(self):
        return not self._rows

    def GetNextRow(self):
        return _Row(self._rows.pop(0))


class _Folder:
    StoreID = "store"

    def __init__(self, mails):
        self.mails = mails
        self.queries = []

    def GetTable(self, filt, _kind):
        self.queries.append(filt)
        # the fake ignores the DASL filter: search() must re-check everything in Python
        return _Table([(m["id"], m["subject"], m["received"], m["sender"], "IPM.Note") for m in self.mails])


class _Item:
    def __init__(self, m):
        self.Body = m["body"]

        class _A:
            Count = 0
        self.Attachments = _A()


class _Session:
    def __init__(self, mails):
        self.folder = _Folder(mails)
        self.by_id = {m["id"]: m for m in mails}
        self.fetches = 0

    def folders(self, _settings):
        return [("Inbox", self.folder)]

    def get_item(self, entry_id, _store):
        self.fetches += 1
        return _Item(self.by_id[entry_id])


NOW = dt.datetime(2026, 11, 5, 12, 0)
MAILS = [
    dict(id="1", subject="RE: NCR-101 bracket", body="Hi, NCR-101 repair is Approved.", sender="A",
         received=NOW - dt.timedelta(days=10)),
    dict(id="2", subject="Question", body="About NCR-1010 and NCR-102: not approved yet", sender="B",
         received=NOW - dt.timedelta(days=3)),
    dict(id="3", subject="NCR-102 concurred", body="ok", sender="C", received=NOW - dt.timedelta(hours=2)),
    dict(id="4", subject="old one NCR-101", body="", sender="D", received=NOW - dt.timedelta(days=900)),
]


def test_bulk_search_assigns_mails_to_ncrs():
    s = OutlookSettings(days_back=730, chunk_size=1)
    sess = _Session(MAILS)
    partials = []
    since = NOW - dt.timedelta(days=730)
    res = om.search(sess, ["NCR-101", "NCR-102"], since, s, partial=lambda r, c: partials.append(c))
    ids = {n: sorted(h.entry_id for h in hits) for n, hits in res.items()}
    assert ids == {"NCR-101": ["1"], "NCR-102": ["2", "3"]}   # NCR-1010 is not NCR-101; mail 4 too old
    by_id = {h.entry_id: h for hits in res.values() for h in hits}
    assert by_id["1"].approved and not by_id["1"].rejected
    assert by_id["2"].rejected and not by_id["2"].approved    # "not approved" is a rejection
    assert by_id["3"].approved
    assert partials == [["NCR-101"], ["NCR-102"]]            # one partial per chunk
    assert len(sess.folder.queries) == 2                     # one query per chunk, not per NCR*folder
    assert sess.fetches == 3                                  # each mail body fetched once only


def test_one_query_for_many_ncrs():
    s = OutlookSettings(chunk_size=40)
    sess = _Session(MAILS)
    om.search(sess, [f"NCR-{i}" for i in range(100, 130)], None, s)
    assert len(sess.folder.queries) == 1
    q = sess.folder.queries[0]
    assert q.startswith("@SQL=(") and q.count("LIKE") == 60  # subject + body for 30 NCRs


def test_incremental_since_filters_old_mail():
    s = OutlookSettings()
    res = om.search(_Session(MAILS), ["NCR-101", "NCR-102"], NOW - dt.timedelta(days=1), s)
    assert {n: [h.entry_id for h in v] for n, v in res.items()} == {"NCR-102": ["3"]}


def test_cancel_stops_search():
    ev = threading.Event()
    ev.set()
    assert om.search(_Session(MAILS), ["NCR-101"], None, OutlookSettings(), cancel=ev) == {}


def test_filter_has_date_and_escapes_quote():
    f = om.build_filter(["A'B"], dt.datetime(2026, 1, 2, 3, 4), True)
    assert "datereceived\" >= '2026-01-02 03:04' AND" in f and "A''B" in f
    assert "datereceived" not in om._strip_date(f)


def test_plan_since_groups_new_and_known():
    s = OutlookSettings(days_back=30)
    checked = {"OLD": NOW - dt.timedelta(hours=5), "NEW": None}
    groups = om.plan_since(["OLD", "NEW"], checked, s, now=NOW)
    assert groups[0] == (NOW - dt.timedelta(hours=5) - om.OVERLAP, ["OLD"])
    assert groups[1] == (NOW - dt.timedelta(days=30), ["NEW"])


def test_cache_merge_save_load(tmp_path, monkeypatch):
    monkeypatch.setattr(om, "CACHE_FILE", tmp_path / "mail.json")
    c = om.MailCache("sig")
    hit = om.MailHit("1", "s", "subj", "A", NOW, "Inbox", approved=True)
    assert c.merge({"NCR-1": [hit]}) == {"NCR-1"}
    assert c.merge({"NCR-1": [hit]}) == set()                # unchanged -> no update signal
    c.mark_checked(["NCR-1"], NOW)
    c.save()
    c2 = om.MailCache.load("sig")
    assert c2.approved("NCR-1") and c2.count("NCR-1") == 1 and c2.checked_at("NCR-1") == NOW
    assert om.MailCache.load("other-settings").count("NCR-1") == 0


def test_match_models_uses_base_number(tmp_path):
    root = tmp_path / "models"
    files = [root / "123A4567.CATPart", root / "sub" / "123A4567-1_rev.CATPart",
             root / "401B2222.CATProduct", root / "999Z0000.CATPart"]
    res = match_models(files, ["123A4567-1"], str(root))
    assert set(res) == {str(files[0]), str(files[1])}
    assert res[str(files[0])] == ["123A4567-1"]
    assert match_models(files, [], str(root)) == {}


def test_norm_path_keeps_unc_on_this_platform():
    assert norm_path("") == ""
    assert norm_path('  "/srv/share/x"  ') == "/srv/share/x"


def test_norm_path_unc_windows(monkeypatch):
    import ncr_panel.utils as u
    monkeypatch.setattr(u, "IS_WINDOWS", True)
    # Qt's folder dialog returns forward slashes for network shares; Explorer needs backslashes
    assert u.norm_path("//vds/Ortak/NCR") == "\\\\vds\\Ortak\\NCR"
    assert u.norm_path("\\\\vds\\Ortak\\NCR\\") == "\\\\vds\\Ortak\\NCR"
    assert u.norm_path("C:/Users/x") == "C:\\Users\\x"
