"""Excel 文件解析服务

支持两种常见格式:
  1) 简单格式 (依据 测试分析001.md excel数据表格式说明):
      多行表头: 项目名称/备注/参数变量/规格上限/规格下限 + 产品数据行
  2) 真实 SC 板格式:
      skip_top_rows=1 跳过顶部 "Program Name:" 行
      接下来 header_rows=5 行为 项目名称/备注/参数变量/上限/下限
      第 fixed_columns_header_row=5 行读取固定列名 (序列号/工单号/名称/测试员/工装编号/结论/时间)
      之后为产品数据

关键处理:
  - 测试项目列若同名且占多列, 自动组合 "项目名称 [参数变量]" 为唯一项名
  - 规格值为 '*' / '-' / 空 视为无限
  - 文本型测试值自动跳过
  - 输出长格式 DataFrame, 便于分组统计
"""
import math
import os
import pandas as pd
from typing import Optional


class ExcelParser:
    """Excel 解析器, 依据产品配置的多行表头规则读取并标准化数据"""

    def __init__(self, product_config: dict):
        """
        excel_config 关键字段:
            sheet_name: 工作表名或序号
            skip_top_rows: 起始跳过的标题行数量 (默认 0)
            header_rows: 测试项目元数据占用行数 (默认 5)
            meta_row_mapping: 项目元数据与行索引映射
                例: {item_name: 0, item_note: 1, param_variable: 2,
                     spec_upper: 3, spec_lower: 4}
            fixed_columns_header_row: 固定列名行号(相对 skip_top_rows 后), 若 <0 则不用该特性
            fixed_columns: 固定列字段映射
                例: {sequence: "序列号", product_model: "产品型号", work_order: "工单号", ...}
        """
        self.config = product_config
        excel_cfg = product_config.get("excel_config", {})
        self.sheet_name = excel_cfg.get("sheet_name", 0)
        self.skip_top_rows = int(excel_cfg.get("skip_top_rows", 0) or 0)
        self.header_rows = int(excel_cfg.get("header_rows", 5) or 5)
        self.meta_row_mapping = excel_cfg.get("meta_row_mapping", {
            "item_name": 0,
            "item_note": 1,
            "param_variable": 2,
            "spec_upper": 3,
            "spec_lower": 4,
        })
        # 关注测试项目白名单: 若非空, 则只解析这些项目 (用于调试或聚焦分析)
        # 项名支持两种写法: 原始项目名(如 "看门狗功能检测_10") 或 带[参数]的唯一名
        self.focus_items = [
            str(x).strip() for x in (product_config.get("test_items") or [])
            if str(x).strip()
        ]
        self.fixed_cols_header_row = int(excel_cfg.get("fixed_columns_header_row", -1) or -1)
        self.fixed_columns = excel_cfg.get("fixed_columns", {
            "sequence": "序列号",
            "product_model": "产品型号",
        })
        # 顶部公共信息 (Program Name 等), 由 _extract_program_info 填充
        self.program_info: dict = {}

        # 规格上下限中表示"无限"的哨兵值 (新格式常用 999999999 表示无上限)
        self.spec_infinity_values = [
            v for v in (excel_cfg.get("spec_infinity_values") or [])
            if v is not None and str(v).strip() != ""
        ]

    def parse(self, file_path: str) -> pd.DataFrame:
        """解析 Excel, 返回长格式 DataFrame (仅数值型测试值)"""
        raw = ExcelParser._read_excel_frame(
            file_path, sheet_name=self.sheet_name, header=None,
        )
        if raw.empty:
            return pd.DataFrame()

        # 0. 提取顶部跳过行中的 Program Name 等公共信息 (保存供后续展示)
        self._extract_program_info(raw)

        # 1. 跳过顶部标题行
        if self.skip_top_rows > 0:
            raw = raw.iloc[self.skip_top_rows:].reset_index(drop=True)

        # 2. 取元数据行
        if len(raw) < self.header_rows + 1:
            return pd.DataFrame()
        header_df = raw.iloc[: self.header_rows].fillna("")

        # 3. 确定数据起始行
        data_start_row = self.header_rows
        # 如果单独一行写固定列名, 读取它并把它当作列名行, 数据从下一行开始
        if self.fixed_cols_header_row >= 0 and self.fixed_cols_header_row < len(raw):
            data_start_row = max(data_start_row, self.fixed_cols_header_row + 1)
            fixed_cols_row = pd.Series(
                [self._strip(v) for v in raw.iloc[self.fixed_cols_header_row].tolist()],
                index=raw.columns,
            )
        else:
            fixed_cols_row = None

        data_df = raw.iloc[data_start_row:].copy()

        # 4. 确定固定列索引
        seq_col_idx = self._find_fixed_col("sequence", header_df, fixed_cols_row)
        model_col_idx = self._find_fixed_col("product_model", header_df, fixed_cols_row)

        # 额外的固定列值 (工单号/结论/时间等) 用于显示, 不影响解析逻辑
        extra_fixed_idx = {}
        for k in self.fixed_columns.keys():
            if k in ("sequence", "product_model"):
                continue
            extra_fixed_idx[k] = self._find_fixed_col(k, header_df, fixed_cols_row)

        # 自动识别"开始时间"和"结束时间"列 (即使配置未声明, 用于 fix_time 不含日期时回退)
        auto_time_cols = {
            "start_time": ["开始时间", "起始时间", "start time", "start_time", "starttime"],
            "end_time":   ["结束时间", "终止时间", "end time", "end_time", "endtime"],
        }
        for k, possible_names in auto_time_cols.items():
            if k in extra_fixed_idx and extra_fixed_idx[k] is not None:
                continue  # 已在配置中声明
            col_idx = self._find_column_by_names(header_df, fixed_cols_row, possible_names)
            if col_idx is not None:
                extra_fixed_idx[k] = col_idx

        # 新格式: 开始时间可能在最后一列且无表头, 内容为时间格式
        if "start_time" not in extra_fixed_idx or extra_fixed_idx["start_time"] is None:
            trailing_time_col = self._detect_trailing_time_column(
                header_df, data_df, excluded={seq_col_idx, model_col_idx}
            )
            if trailing_time_col is not None:
                extra_fixed_idx["start_time"] = trailing_time_col

        # 5. 收集测试项目列 (跳过固定列)
        excluded_cols = {
            c for c in (
                [seq_col_idx, model_col_idx] + list(extra_fixed_idx.values())
            ) if c is not None
        }
        test_item_cols = self._collect_test_items(header_df, excluded_cols)

        if not test_item_cols:
            raise ValueError("未在表头中识别到测试项目列, 请检查配置")

        # 6. 逐行解析产品数据, 构造长格式记录
        records = []
        for _, row in data_df.iterrows():
            seq_val = row.iloc[seq_col_idx] if seq_col_idx is not None else None
            if not self._has_value(seq_val):
                continue
            model_val = row.iloc[model_col_idx] if model_col_idx is not None else ""

            fixed_extra = {
                k: (row.iloc[i] if i is not None and i < len(row) else "")
                for k, i in extra_fixed_idx.items()
            }

            for item in test_item_cols:
                col_idx = item["col_index"]
                raw_val = row.iloc[col_idx] if col_idx < len(row) else None

                numeric_val = self._try_numeric(raw_val)
                if numeric_val is None:
                    continue  # 跳过文本型

                rec = {
                    "sequence": str(seq_val),
                    "product_model": str(model_val) if model_val else "",
                    "item_name": item["item_name"],
                    "item_value": numeric_val,
                    "item_note": item.get("item_note", ""),
                    "param_variable": item.get("param_variable", ""),
                    "spec_upper": item.get("spec_upper"),
                    "spec_lower": item.get("spec_lower"),
                }
                for ek, ev in fixed_extra.items():
                    rec["fix_" + ek] = "" if ev is None or (isinstance(ev, float) and math.isnan(ev)) else str(ev)
                records.append(rec)

        return pd.DataFrame(records)

    # ------------------------------------------------------------------ #
    #  内部辅助
    # ------------------------------------------------------------------ #
    def _find_fixed_col(self, field_key: str, header_df: pd.DataFrame,
                        fixed_cols_row: Optional[pd.Series]) -> Optional[int]:
        """查找固定列(序列号/工单号等)的列索引"""
        target_name = self.fixed_columns.get(field_key, "")
        if not target_name:
            return None

        # 优先从 fixed_cols_row 匹配
        if fixed_cols_row is not None:
            for i, v in enumerate(fixed_cols_row.astype(str).tolist()):
                if v.strip() == target_name.strip():
                    return i

        # 回退: 用 项目名称 行匹配 (简单格式)
        name_row_idx = self.meta_row_mapping.get("item_name", 0)
        if 0 <= name_row_idx < len(header_df):
            name_row = header_df.iloc[name_row_idx].astype(str).str.strip()
            matches = name_row[name_row == target_name.strip()].index.tolist()
            if matches:
                return int(matches[0])
        return None

    def _find_column_by_names(self, header_df: pd.DataFrame,
                              fixed_cols_row: Optional[pd.Series],
                              possible_names: list) -> Optional[int]:
        """根据一组可能的列名在表头中查找列索引 (用于自动识别开始/结束时间等)"""
        norm_names = {n.strip().lower() for n in possible_names if n}

        # 1. 优先从 fixed_cols_row 匹配 (固定列名行)
        if fixed_cols_row is not None:
            for i, v in enumerate(fixed_cols_row.astype(str).tolist()):
                if v.strip().lower() in norm_names:
                    return i

        # 2. 回退: 用 项目名称 行匹配 (简单格式)
        name_row_idx = self.meta_row_mapping.get("item_name", 0)
        if 0 <= name_row_idx < len(header_df):
            name_row = header_df.iloc[name_row_idx].astype(str).str.strip()
            for i, v in enumerate(name_row.tolist()):
                if v.lower() in norm_names:
                    return i
        return None

    @staticmethod
    def _detect_trailing_time_column(header_df: pd.DataFrame,
                                     data_df: pd.DataFrame,
                                     excluded: set) -> Optional[int]:
        """检测最后一列是否为无表头的时间列 (新格式开始时间)

        条件: 该列在表头中无名称, 且数据行中包含时间格式字符串 (如 2026-09-01-16:41)
        """
        import re
        name_row_idx = 0
        if name_row_idx >= len(header_df):
            return None
        name_row = header_df.iloc[name_row_idx].astype(str).str.strip()
        ncols = len(name_row)
        # 从最后一列向前查找无表头的列
        for col_idx in range(ncols - 1, -1, -1):
            if col_idx in excluded:
                continue
            header = name_row.iloc[col_idx]
            if header and header != "nan":
                # 有表头, 不是目标列 (新格式时间列无表头)
                # 但仍可能是普通测试项目列, 继续向前找
                continue
            # 检查该列数据是否含时间格式
            time_pattern = re.compile(
                r"\d{4}[-/]\d{1,2}[-/]\d{1,2}([ T]\d{1,2}:\d{2})?"
            )
            sample = data_df.iloc[:, col_idx].dropna().head(5)
            for v in sample:
                s = str(v).strip()
                if s and time_pattern.search(s):
                    return col_idx
            # 该列无表头但也无时间值, 继续向前
        return None

    def _collect_test_items(self, header_df: pd.DataFrame, excluded: set) -> list:
        """从元数据行提取测试项目 (自动组合同名+参数变量为唯一项名)"""
        name_row_idx = self.meta_row_mapping.get("item_name", 0)
        if name_row_idx >= len(header_df):
            return []
        # 对每个单元格去空格和引号
        raw_names = [self._strip(v) for v in header_df.iloc[name_row_idx].tolist()]
        name_series = pd.Series(raw_names, index=header_df.columns)

        # 记录 "项目名称" 列使用次数, 判断是否多列共享同一项目名 (忽略空值/星号)
        valid_counts = pd.Series(
            [n for n in name_series.tolist() if n and n != "*"]
        ).value_counts()

        # 行标签列名 (如 "ITEM" / "项目名称"), 这些列仅作为表头标签, 不是测试项目
        label_names = {"item", "项目名称", "测试项目", "item_name", "item name"}

        items = []
        for col_idx in range(len(name_series)):
            if col_idx in excluded:
                continue
            raw_item_name = name_series.iloc[col_idx]
            if not raw_item_name or raw_item_name == "*":
                continue
            # 跳过行标签列
            if raw_item_name.strip().lower() in label_names:
                continue

            param = self._cell(header_df, self.meta_row_mapping.get("param_variable"), col_idx)
            # 如果同一个名称占用多列，就拼接 [参数变量] 使名称唯一
            if valid_counts.get(raw_item_name, 1) > 1 and param:
                final_name = f"{raw_item_name} [{param}]"
            else:
                final_name = raw_item_name

            # 白名单过滤: 若配置了 test_items, 只保留匹配项
            # 匹配规则: final_name 完全匹配, 或 raw_item_name 完全匹配 (便于用原始名指定整组)
            if self.focus_items:
                if final_name not in self.focus_items and raw_item_name not in self.focus_items:
                    continue

            item = {
                "col_index": col_idx,
                "item_name": final_name,
                "item_note": self._cell(header_df, self.meta_row_mapping.get("item_note"), col_idx),
                "param_variable": param,
            }
            item["spec_upper"] = self._try_spec_numeric(
                self._cell(header_df, self.meta_row_mapping.get("spec_upper"), col_idx),
                self.spec_infinity_values,
            )
            item["spec_lower"] = self._try_spec_numeric(
                self._cell(header_df, self.meta_row_mapping.get("spec_lower"), col_idx),
                self.spec_infinity_values,
            )
            items.append(item)

        return items

    def _extract_program_info(self, raw: pd.DataFrame) -> None:
        """从顶部跳过的标题行提取公共信息 (Program Name 等)

        支持两种写法:
          1) 同一单元格内 "Key: Value" / "Key：Value"
          2) Key 单独占一单元格 (以 : 或 ：结尾), 值在右侧相邻单元格
        """
        info = {}
        if self.skip_top_rows <= 0 or raw.empty:
            self.program_info = info
            return
        top = raw.iloc[: self.skip_top_rows]
        for r in range(top.shape[0]):
            for c in range(top.shape[1]):
                v = top.iat[r, c]
                s = self._strip(v)
                if not s or s == "nan":
                    continue
                # 写法 1: 同一单元格内 "Key: Value" (仅当冒号后确实有值时才匹配)
                matched = False
                for sep in (":", "："):
                    if sep in s:
                        k, _, val = s.partition(sep)
                        k = k.strip()
                        val = val.strip()
                        if k and val and k.lower() != "nan":
                            info[k] = val
                            matched = True
                        break
                if matched:
                    continue
                # 写法 2: Key 单独占单元格 (以 : 或 ：结尾), 值在右侧
                if s.endswith(":") or s.endswith("："):
                    k = s.rstrip(":：").strip()
                    if k and c + 1 < top.shape[1]:
                        val = self._strip(top.iat[r, c + 1])
                        if val and val.lower() != "nan":
                            info[k] = val
        self.program_info = info

    @staticmethod
    def _strip(value) -> str:
        """将值转字符串, 去除首尾空格和引号 (Excel 中常见 "'内容'" 形式)"""
        if value is None:
            return ""
        s = str(value).strip()
        while len(s) >= 2 and s[0] == s[-1] and s[0] in ("'", '"'):
            s = s[1:-1]
        return "" if s in ("nan", "None") else s

    @staticmethod
    def _cell(header_df: pd.DataFrame, row_idx: Optional[int], col_idx: int) -> str:
        """安全读取表头单元格"""
        if row_idx is None or row_idx >= len(header_df) or col_idx >= len(header_df.columns):
            return ""
        return ExcelParser._strip(header_df.iloc[row_idx, col_idx])

    @staticmethod
    def _has_value(v) -> bool:
        """单元格是否有非空内容"""
        if v is None:
            return False
        if isinstance(v, float) and math.isnan(v):
            return False
        s = ExcelParser._strip(v)
        return s != ""

    @staticmethod
    def _try_numeric(val) -> Optional[float]:
        """尝试将值转为 float, 保留原始精度以避免浮点误差"""
        if val is None:
            return None
        if isinstance(val, bool):
            return None
        if isinstance(val, (int, float)):
            if isinstance(val, float) and math.isnan(val):
                return None
            return float(val)
        s = ExcelParser._strip(val)
        if not s or s in ("-", "—", "/", "*"):
            return None
        try:
            f = float(s)
        except (ValueError, TypeError):
            return None
        if "." in s:
            decimals = len(s.split(".")[-1])
            return round(f, decimals)
        return f

    @classmethod
    def _try_spec_numeric(cls, val_str: str,
                          infinity_values: Optional[list] = None) -> Optional[float]:
        """规格上下限转数值: '*' / '-' / 哨兵值 视为无限 (None), 保留原始精度

        Args:
            val_str: 原始单元格字符串
            infinity_values: 额外的无限哨兵值列表 (如 [999999999]), 匹配时返回 None
        """
        s = cls._strip(val_str)
        if not s or s in ("*", "-", "—", "/"):
            return None
        try:
            f = float(s)
        except (ValueError, TypeError):
            return None
        # 匹配无限哨兵值 (字符串或数值比较)
        if infinity_values:
            for iv in infinity_values:
                try:
                    if f == float(iv) or s == str(iv).strip():
                        return None
                except (ValueError, TypeError):
                    if s == str(iv).strip():
                        return None
        if "." in s:
            decimals = len(s.split(".")[-1])
            return round(f, decimals)
        return f

    @staticmethod
    def get_sheet_names(file_path: str) -> list:
        """获取 Excel 文件中所有工作表名称 (文本格式返回 ["Sheet1"])"""
        fmt = ExcelParser._detect_format(file_path)
        if fmt in ("xlsx", "xls"):
            engine = "openpyxl" if fmt == "xlsx" else "xlrd"
            return pd.ExcelFile(file_path, engine=engine).sheet_names
        return ["Sheet1"]

    # ------------------------------------------------------------------ #
    #  自动检测配置
    # ------------------------------------------------------------------ #
    @staticmethod
    def auto_detect_config(file_path: str) -> dict:
        """从 Excel 文件结构自动检测解析配置

        当没有可用产品配置时, 通过扫描文件内容推断:
          - skip_top_rows: 标题行数 (如 "Program Name:" 行)
          - header_rows / meta_row_mapping: 测试项目元数据行
          - fixed_columns_header_row: 固定列名所在行 (相对 skip 后)
          - fixed_columns: 固定列字段映射 (序列号/工单号/时间等)

        检测策略:
          1) 扫描前几行查找 "Program Name" / "程序名" -> skip_top_rows
          2) 在 skip 后的行中查找 "规格上限"/"规格下限"/"项目名称" 等标签 -> meta_row_mapping
          3) 从 header_rows 之后扫描 "序列号"/"时间"/"工单号" 等 -> fixed_columns
        """
        raw = ExcelParser._read_excel_frame(
            file_path, sheet_name=0, header=None,
        )
        if raw.empty:
            return ExcelParser._default_config()

        scan_max = min(25, len(raw))

        # 固定列关键词: (字段名, [可能的列名])
        FIXED_COL_KW = [
            ("sequence",      ["序列号", "序号", "sequence", "seq", "sn", "sn003"]),
            ("product_model", ["产品型号", "型号", "product_model", "model"]),
            ("work_order",    ["工单号", "工单", "work_order", "work order", "equcode"]),
            ("tester",        ["测试员", "测试人", "tester", "操作员"]),
            ("fixture",       ["工装编号", "工装号", "工装", "fixture"]),
            ("result",        ["结论", "结果", "result"]),
            ("time",          ["时间", "测试时间", "time", "test time"]),
            ("start_time",    ["开始时间", "起始时间", "start time", "start_time", "starttime"]),
            ("end_time",      ["结束时间", "终止时间", "end time", "end_time", "endtime"]),
        ]
        # 元数据行标签关键词 (同时支持中文和英文标签)
        META_KW = {
            "spec_upper": {"规格上限", "上限", "spec_upper", "spec upper", "spec max", "usl", "upper"},
            "spec_lower": {"规格下限", "下限", "spec_lower", "spec lower", "spec min", "lsl", "lower"},
            "item_name":  {"项目名称", "测试项目", "item_name", "item name", "item"},
            "param_variable": {"参数变量", "参数", "param_variable", "param", "variable", "unit"},
            "item_note":  {"备注", "说明", "item_note", "note", "项目备注"},
        }

        # 1. 检测 skip_top_rows: 查找标题行 (Program Name / Model / EQUCODE 等)
        skip_top_rows = 0
        for r in range(min(5, scan_max)):
            row_text = " ".join(ExcelParser._strip(v) for v in raw.iloc[r].tolist()).lower()
            if ("program name" in row_text or "程序名" in row_text
                    or "程序名称" in row_text or "model:" in row_text
                    or "equcode" in row_text):
                skip_top_rows = r + 1
                break

        # 2. 检测 meta_row_mapping: 在 skip 后的行中查找标签关键词
        meta_row_map = {}
        meta_search_end = min(skip_top_rows + 10, scan_max)
        for r in range(skip_top_rows, meta_search_end):
            # 只扫描前几列 (标签通常在左侧)
            for c in range(min(8, raw.shape[1])):
                s = ExcelParser._strip(raw.iloc[r, c]).lower()
                if not s:
                    continue
                for meta_key, kw_set in META_KW.items():
                    if s in kw_set and meta_key not in meta_row_map:
                        meta_row_map[meta_key] = r - skip_top_rows

        # 补全默认值
        defaults = {
            "item_name": 0, "item_note": 1, "param_variable": 2,
            "spec_upper": 3, "spec_lower": 4,
        }
        for k, v in defaults.items():
            if k not in meta_row_map:
                meta_row_map[k] = v

        # header_rows: 元数据行数 (取最大偏移 + 1; 新格式可能只有 4 行无备注行)
        header_rows = max(meta_row_map.values()) + 1

        # 3. 检测 fixed_columns_header_row 和 fixed_columns
        # 搜索范围: 从 skip 后第一行 (ITEM/项目名称行, 新格式固定列可能在此行)
        #           到 header_rows 之后几行 (旧格式固定列名单独占一行)
        fixed_header_row_abs = -1
        fixed_cols = {}
        search_start = skip_top_rows
        search_end = min(skip_top_rows + header_rows + 6, scan_max)
        for r in range(search_start, search_end):
            row_strs = [ExcelParser._strip(v) for v in raw.iloc[r].tolist()]
            row_lower = [s.lower() for s in row_strs]
            found_in_row = False
            for field, keywords in FIXED_COL_KW:
                kw_lower = [kw.lower() for kw in keywords]
                for c, s_low in enumerate(row_lower):
                    if s_low and s_low in kw_lower:
                        if fixed_header_row_abs < 0:
                            fixed_header_row_abs = r
                        if field not in fixed_cols:
                            fixed_cols[field] = row_strs[c]
                        found_in_row = True
                        break
            # 找到足够多的固定列就停止
            if found_in_row and len(fixed_cols) >= 2:
                break

        # 若固定列出现在元数据行内 (如新格式 SN/Result 与 ITEM 同行),
        # 则不需要单独的 fixed_columns_header_row, 设为 -1 让解析器从 item_name 行匹配
        if (fixed_header_row_abs >= 0
                and fixed_header_row_abs < skip_top_rows + header_rows):
            fixed_cols_header_row = -1
        else:
            fixed_cols_header_row = (fixed_header_row_abs - skip_top_rows
                                     if fixed_header_row_abs >= 0 else -1)

        # 4. 检测规格无限哨兵值 (如新格式用 999999999 表示无上限)
        spec_infinity_values = []
        for meta_key in ("spec_upper", "spec_lower"):
            row_abs = skip_top_rows + meta_row_map.get(meta_key, 0)
            if row_abs < len(raw):
                for c in range(min(raw.shape[1], 50)):
                    v = ExcelParser._strip(raw.iloc[row_abs, c])
                    if not v:
                        continue
                    try:
                        fv = float(v)
                    except (ValueError, TypeError):
                        continue
                    # 极大值 (>= 1e8) 视为无限哨兵
                    if abs(fv) >= 1e8 and fv not in spec_infinity_values:
                        spec_infinity_values.append(fv)

        # 5. 构建配置
        excel_cfg = {
            "sheet_name": 0,
            "skip_top_rows": skip_top_rows,
            "header_rows": header_rows,
            "meta_row_mapping": meta_row_map,
            "fixed_columns_header_row": fixed_cols_header_row,
            "fixed_columns": fixed_cols if fixed_cols else {
                "sequence": "序列号",
                "product_model": "产品型号",
            },
            "spec_infinity_values": spec_infinity_values,
        }

        return {
            "category_name": "自动检测配置",
            "excel_config": excel_cfg,
            "test_items": [],
            "merge_items": [],
            "chart_config": {
                "title": "测试分析曲线图",
                "y_axis_name": "测试值",
            },
        }

    @staticmethod
    def _detect_format(file_path: str) -> str:
        """根据文件内容(魔数)检测真实格式, 而非仅靠扩展名

        产线导出的 "xls" 常见真实形态:
          - OLE2/BIFF 二进制 (.xls) -> xlrd
          - OOXML zip 容器 (.xlsx)   -> openpyxl
          - 制表符分隔文本 (伪装 .xls)-> read_csv(sep='\t')
          - HTML 表格 (伪装 .xls)    -> read_html

        返回: 'xlsx' | 'xls' | 'html' | 'tsv' | 'csv'
        """
        try:
            with open(file_path, "rb") as f:
                head = f.read(8)
        except OSError:
            return "xlsx"
        if head[:4] == b"PK\x03\x04":
            return "xlsx"
        if head[:8] == b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1":
            return "xls"
        sample = head.decode("latin-1", errors="ignore").lstrip("\ufeff").lstrip()
        if sample.startswith("<"):
            return "html"
        try:
            with open(file_path, "rb") as f:
                chunk = f.read(4096).decode("latin-1", errors="ignore")
        except OSError:
            chunk = ""
        if chunk.count("\t") >= chunk.count(","):
            return "tsv"
        return "csv"

    @staticmethod
    def _read_excel_engine(file_path: str) -> str:
        """仅对真正的 Excel 文件返回 read_excel 引擎名 (兼容旧调用)"""
        fmt = ExcelParser._detect_format(file_path)
        if fmt == "xls":
            return "xlrd"
        return "openpyxl"

    @staticmethod
    def _read_excel_frame(file_path: str, sheet_name=0, header=None) -> pd.DataFrame:
        """读取为 DataFrame, 按真实格式自动适配 (兼容伪装成 .xls 的文本/HTML)"""
        fmt = ExcelParser._detect_format(file_path)
        if fmt == "xlsx":
            return pd.read_excel(file_path, sheet_name=sheet_name, header=header, engine="openpyxl")
        if fmt == "xls":
            return pd.read_excel(file_path, sheet_name=sheet_name, header=header, engine="xlrd")
        if fmt == "html":
            tables = pd.read_html(file_path, header=header)
            return tables[0] if tables else pd.DataFrame()
        sep = "\t" if fmt == "tsv" else ","
        # 中文产线文件常为 GBK/GB18030 编码, 逐个尝试常见编码
        last_exc = None
        for enc in ("utf-8-sig", "gb18030", "latin-1"):
            try:
                return pd.read_csv(file_path, sep=sep, header=header, dtype=str, encoding=enc)
            except UnicodeDecodeError as exc:
                last_exc = exc
                continue
        raise last_exc

    @staticmethod
    def _default_config() -> dict:
        """返回默认配置模板 (与 ConfigManager.default_config 一致)"""
        return {
            "category_name": "",
            "excel_config": {
                "sheet_name": 0,
                "skip_top_rows": 0,
                "header_rows": 5,
                "meta_row_mapping": {
                    "item_name": 0,
                    "item_note": 1,
                    "param_variable": 2,
                    "spec_upper": 3,
                    "spec_lower": 4,
                },
                "fixed_columns_header_row": -1,
                "fixed_columns": {
                    "sequence": "序列号",
                    "product_model": "产品型号",
                },
                "spec_infinity_values": [],
            },
            "test_items": [],
            "merge_items": [],
            "chart_config": {
                "title": "测试分析曲线图",
                "y_axis_name": "测试值",
            },
        }
