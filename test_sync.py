# -*- coding: utf-8 -*-
"""
sync.py 的自测，不联网 —— 用假的 requests 顶替飞书接口。

    python test_sync.py
"""

import json
import os
import sys
import tempfile
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import sync


HEADERS = [
    "报表类型", "统计周期", "用户ID", "账户等级", "商户名称", "直签人",
    "代理商用户ID", "代理商名称", "站点", "接入模式",
    "交易金额（美元）", "环比增长金额（美元）", "金额环比",
    "交易笔数", "环比增长笔数", "笔数环比",
]

# 多维表格的字段类型：ID 和周期是文本，金额笔数是数字
TABLE_FIELDS = dict(
    {h: sync.FT_TEXT for h in HEADERS},
    **{h: sync.FT_NUMBER for h in sync.NUMBER_COLUMNS}
)


def make_xlsx(path, rows, headers=HEADERS, title_row=True):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    if title_row:
        ws.append(["出单统计报表"])
    ws.append(headers)
    for r in rows:
        ws.append(r)
    wb.save(path)
    return path


def row(period, uid, name="商户A", amt=100.0, cnt=5):
    return ["日报", period, uid, "V4", name, "赵娜", "", "",
            "www.a.com", "Element", amt, 0, 0, cnt, 0, 0]


class FakeFeishu:
    """按 record_id 存记录，模拟 upsert 语义"""

    def __init__(self, fields=None, existing=None, support_search=True):
        self.fields = fields if fields is not None else dict(TABLE_FIELDS)
        self.records = dict(existing or {})
        self.support_search = support_search
        self.created, self.updated = [], []
        self.calls = []
        self._next = 100

    def __call__(self, method, url, **kw):
        self.calls.append((method, url))
        body = kw.get("json") or {}

        if "tenant_access_token" in url:
            return self._ok({"tenant_access_token": "t-fake", "expire": 7200})

        if url.endswith("/fields"):
            return self._ok({
                "items": [{"field_name": n, "type": t} for n, t in self.fields.items()],
                "has_more": False,
            })

        if url.endswith("/records/search"):
            if not self.support_search:
                return self._err(1254005, "search not supported")
            wanted = set()
            for c in body.get("filter", {}).get("conditions", []):
                v = c["value"]
                wanted.add(v[1] if v[0] == "ExactDate" else v[0])
            items = [
                {"record_id": rid, "fields": f}
                for rid, f in self.records.items()
                if f.get("统计周期") in wanted
            ]
            return self._ok({"items": items, "has_more": False})

        if url.endswith("/records"):
            items = [{"record_id": rid, "fields": f} for rid, f in self.records.items()]
            return self._ok({"items": items, "has_more": False})

        if url.endswith("/records/batch_create"):
            out = []
            for r in body["records"]:
                rid = f"rec{self._next}"
                self._next += 1
                self.records[rid] = r["fields"]
                self.created.append(r["fields"])
                out.append({"record_id": rid, "fields": r["fields"]})
            return self._ok({"records": out})

        if url.endswith("/records/batch_update"):
            for r in body["records"]:
                self.records[r["record_id"]] = r["fields"]
                self.updated.append((r["record_id"], r["fields"]))
            return self._ok({"records": body["records"]})

        raise AssertionError(f"预期外的请求: {method} {url}")

    @staticmethod
    def _ok(data):
        return mock.Mock(json=lambda: {"code": 0, "msg": "success", "data": data})

    @staticmethod
    def _err(code, msg):
        return mock.Mock(json=lambda: {"code": code, "msg": msg})


class TestParsing(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def path(self, name="r.xlsx"):
        return os.path.join(self.tmp, name)

    def test_读出基本字段并按类型转换(self):
        p = make_xlsx(self.path(), [row("2026-08-06", "1828351428978479106")])
        recs = sync.read_report(p)
        self.assertEqual(len(recs), 1)
        r = recs[0]
        self.assertEqual(r["用户ID"], "1828351428978479106")     # 字符串，未被压成 float
        self.assertEqual(r["统计周期"], "2026-08-06")
        self.assertIsInstance(r["交易金额（美元）"], float)

    def test_表头前有标题行也能定位(self):
        p = make_xlsx(self.path(), [row("2026-08-06", "111")], title_row=True)
        self.assertEqual(len(sync.read_report(p)), 1)

    def test_跳过合计行和空行(self):
        rows = [
            row("2026-08-06", "111"),
            [None, None, None, None, "合计"] + [None] * 11,   # 无用户ID
            [None] * 16,
        ]
        p = make_xlsx(self.path(), rows)
        self.assertEqual(len(sync.read_report(p)), 1)

    def test_日期斜杠格式归一化(self):
        p = make_xlsx(self.path(), [row("2026/08/06", "111")])
        self.assertEqual(sync.read_report(p)[0]["统计周期"], "2026-08-06")

    def test_百分号字符串转小数(self):
        self.assertAlmostEqual(sync.to_float("-8.84%"), -0.0884)
        self.assertAlmostEqual(sync.to_float("1.78%"), 0.0178)
        self.assertEqual(sync.to_float("62,557.00"), 62557.0)
        self.assertIsNone(sync.to_float(""))
        self.assertIsNone(sync.to_float("-"))

    def test_用户ID被存成数字时中止而不是写错数据(self):
        p = make_xlsx(self.path(), [row("2026-08-06", 1828351428978479107)])
        with self.assertRaises(sync.SyncError):
            sync.read_report(p)

    def test_小的数字ID不误报(self):
        p = make_xlsx(self.path(), [row("2026-08-06", 12345)])
        self.assertEqual(sync.read_report(p)[0]["用户ID"], "12345")

    def test_报表多出未知列不报错(self):
        p = make_xlsx(self.path(), [row("2026-08-06", "111") + ["新东西"]],
                      headers=HEADERS + ["新增列X"])
        self.assertEqual(sync.read_report(p)[0]["新增列X"], "新东西")

    def test_找不到表头时报错(self):
        p = make_xlsx(self.path(), [["a", "b"]], headers=["随便", "什么"], title_row=False)
        with self.assertRaises(sync.SyncError):
            sync.read_report(p)

    def test_取目录里最新的文件(self):
        import time as _t
        a = make_xlsx(self.path("a.xlsx"), [row("2026-08-05", "111")])
        _t.sleep(0.01)
        b = make_xlsx(self.path("b.xlsx"), [row("2026-08-06", "111")])
        os.utime(b, (_t.time() + 10, _t.time() + 10))
        self.assertEqual(sync.pick_latest_file(self.tmp), b)
        open(os.path.join(self.tmp, "~$tmp.xlsx"), "w").close()   # Excel 锁文件要忽略
        self.assertEqual(sync.pick_latest_file(self.tmp), b)


class TestBuildFields(unittest.TestCase):

    def test_数字字段写数字文本字段写字符串(self):
        rec = {"用户ID": "1828351428978479106", "交易笔数": 1084.0, "统计周期": "2026-08-06"}
        out = sync.build_fields(rec, TABLE_FIELDS)
        self.assertEqual(out["用户ID"], "1828351428978479106")
        self.assertEqual(out["交易笔数"], 1084.0)
        self.assertEqual(out["统计周期"], "2026-08-06")

    def test_整数不写成带小数点的字符串(self):
        out = sync.build_fields({"账户等级": 4.0}, {"账户等级": sync.FT_TEXT})
        self.assertEqual(out["账户等级"], "4")

    def test_表里没有的列被丢掉(self):
        out = sync.build_fields({"用户ID": "1", "报表里独有的列": "x"},
                                {"用户ID": sync.FT_TEXT})
        self.assertEqual(out, {"用户ID": "1"})

    def test_日期字段转毫秒时间戳(self):
        out = sync.build_fields({"统计周期": "2026-08-06"}, {"统计周期": sync.FT_DATE})
        self.assertIsInstance(out["统计周期"], int)
        self.assertEqual(sync.to_date_str(
            __import__("datetime").datetime.fromtimestamp(out["统计周期"] / 1000)
        ), "2026-08-06")

    def test_单选字段空值不写(self):
        out = sync.build_fields({"接入模式": ""}, {"接入模式": sync.FT_SINGLE_SELECT})
        self.assertNotIn("接入模式", out)


class TestSync(unittest.TestCase):
    """跑通 main()：解析 → 读已有 → 新增/更新"""

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        cfg = os.path.join(self.tmp, "config.json")
        with open(cfg, "w", encoding="utf-8") as f:
            json.dump({"feishu": {
                "app_id": "cli_x", "app_secret": "s",
                "bitable": {"app_token": "bas_x", "table_id": "tbl_x"},
            }}, f)
        self.cfg = cfg

    def run_sync(self, rows, fake, extra_args=()):
        p = make_xlsx(os.path.join(self.tmp, "r.xlsx"), rows)
        argv = ["sync.py", "--file", p, "--config", self.cfg, *extra_args]
        with mock.patch.object(sync.requests, "request", side_effect=fake), \
             mock.patch.object(sys, "argv", argv), \
             mock.patch.object(sync.time, "sleep"):
            return sync.main()

    def test_空表全部新增(self):
        fake = FakeFeishu()
        rc = self.run_sync([row("2026-08-06", "111"), row("2026-08-06", "222")], fake)
        self.assertEqual(rc, 0)
        self.assertEqual(len(fake.created), 2)
        self.assertEqual(len(fake.updated), 0)

    def test_重跑同一天只更新不重复新增(self):
        fake = FakeFeishu(existing={
            "rec1": {"统计周期": "2026-08-06", "用户ID": "111", "交易金额（美元）": 50.0},
        })
        rc = self.run_sync([row("2026-08-06", "111", amt=999.0),
                            row("2026-08-06", "222")], fake)
        self.assertEqual(rc, 0)
        self.assertEqual(len(fake.created), 1)                  # 只有 222 是新的
        self.assertEqual(len(fake.updated), 1)                  # 111 被更新
        self.assertEqual(fake.updated[0][0], "rec1")
        self.assertEqual(fake.updated[0][1]["交易金额（美元）"], 999.0)
        self.assertEqual(len(fake.records), 2)                  # 没变成 3 条

    def test_不同日期同一商户各存一条(self):
        fake = FakeFeishu()
        self.run_sync([row("2026-08-05", "111"), row("2026-08-06", "111")], fake)
        self.assertEqual(len(fake.created), 2)

    def test_报表内重复key只写最后一条(self):
        fake = FakeFeishu()
        self.run_sync([row("2026-08-06", "111", name="旧"),
                       row("2026-08-06", "111", name="新")], fake)
        self.assertEqual(len(fake.created), 1)
        self.assertEqual(fake.created[0]["商户名称"], "新")

    def test_search不可用时回退全量(self):
        fake = FakeFeishu(support_search=False, existing={
            "rec1": {"统计周期": "2026-08-06", "用户ID": "111"},
        })
        rc = self.run_sync([row("2026-08-06", "111")], fake)
        self.assertEqual(rc, 0)
        self.assertEqual(len(fake.updated), 1)                  # 回退后仍能匹配上
        self.assertTrue(any(u.endswith("/records") for _, u in fake.calls))

    def test_缺唯一键字段时中止(self):
        fields = {k: v for k, v in TABLE_FIELDS.items() if k != "用户ID"}
        fake = FakeFeishu(fields=fields)
        rc = self.run_sync([row("2026-08-06", "111")], fake)
        self.assertEqual(rc, 1)
        self.assertEqual(len(fake.created), 0)                  # 什么都没写

    def test_表里缺普通字段时跳过该列继续写(self):
        fields = {k: v for k, v in TABLE_FIELDS.items() if k != "账户等级"}
        fake = FakeFeishu(fields=fields)
        rc = self.run_sync([row("2026-08-06", "111")], fake)
        self.assertEqual(rc, 0)
        self.assertNotIn("账户等级", fake.created[0])
        self.assertIn("商户名称", fake.created[0])

    def test_date参数只同步指定周期(self):
        fake = FakeFeishu()
        self.run_sync([row("2026-08-05", "111"), row("2026-08-06", "222")],
                      fake, extra_args=("--date", "2026-08-06"))
        self.assertEqual(len(fake.created), 1)
        self.assertEqual(fake.created[0]["用户ID"], "222")

    def test_缺配置时报错退出(self):
        empty = os.path.join(self.tmp, "empty.json")
        with open(empty, "w") as f:
            json.dump({}, f)
        p = make_xlsx(os.path.join(self.tmp, "r.xlsx"), [row("2026-08-06", "111")])
        env = {k: "" for k in ("FEISHU_APP_ID", "FEISHU_APP_SECRET",
                               "BITABLE_APP_TOKEN", "BITABLE_TABLE_ID")}
        with mock.patch.dict(os.environ, env, clear=False), \
             mock.patch.object(sys, "argv", ["sync.py", "--file", p, "--config", empty]):
            self.assertEqual(sync.main(), 1)

    def test_dry_run不发任何请求(self):
        fake = FakeFeishu()
        rc = self.run_sync([row("2026-08-06", "111")], fake, extra_args=("--dry-run",))
        self.assertEqual(rc, 0)
        self.assertEqual(fake.calls, [])


class TestConfig(unittest.TestCase):

    def test_环境变量优先于配置文件(self):
        tmp = tempfile.mkdtemp()
        cfg = os.path.join(tmp, "c.json")
        with open(cfg, "w", encoding="utf-8") as f:
            json.dump({"feishu": {"app_id": "from_file",
                                  "bitable": {"app_token": "tok_file"}}}, f)
        with mock.patch.dict(os.environ, {"FEISHU_APP_ID": "from_env"}, clear=False):
            c = sync.load_config(cfg)
        self.assertEqual(c["app_id"], "from_env")
        self.assertEqual(c["app_token"], "tok_file")


if __name__ == "__main__":
    unittest.main(verbosity=2)
