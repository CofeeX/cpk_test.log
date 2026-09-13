"""统计分析服务

使用 Pandas 对长格式测试数据进行统计分析: 均值/标准差/最值/分位数/合格率等。
输出包含规格上下限, 供前端 ECharts 绘制参考线 (markLine)。

输入 DataFrame 标准列:
    sequence | product_model | item_name | item_value(数值)
    | item_note | param_variable | spec_upper | spec_lower
"""
import math
import re
import pandas as pd
from datetime import datetime
from typing import Optional


class Analyzer:
    """测试数据统计分析器"""

    # 不参与分析的项目名关键词 (项目名包含任一关键词即跳过)
    EXCLUDE_ITEM_KEYWORDS = ["CAN_报文定义"]
    # 数字量/状态量参数标记: 参数变量为 D (或中文标记) 的项目不参与统计分析
    EXCLUDE_PARAM_KEYWORDS = ("数字量", "状态量", "数字", "状态", "digital")

    @staticmethod
    def _should_exclude_param(param_variable) -> bool:
        """判断参数变量是否为数字量/状态量 (不参与统计分析)

        规则: 参数变量精确为 "D" / "DIGITAL" (产线数字量标记),
              或包含 中文/英文 关键标记词。
        """
        if param_variable is None:
            return False
        p = str(param_variable).strip()
        if not p:
            return False
        if p.upper() in ("D", "DIGITAL", "DISCRETE"):
            return True
        pl = p.lower()
        return any(kw in pl for kw in Analyzer.EXCLUDE_PARAM_KEYWORDS)

    @staticmethod
    def _should_exclude_item(item_name, param_variable=None) -> bool:
        """判断项目是否应被排除 (不参与分析): 名称关键词 或 数字量/状态量参数"""
        if item_name:
            s = str(item_name)
            if any(kw in s for kw in Analyzer.EXCLUDE_ITEM_KEYWORDS):
                return True
        return Analyzer._should_exclude_param(param_variable)

    @staticmethod
    def _item_excluded(df: pd.DataFrame, item_name) -> bool:
        """基于数据内容判断项目是否应被排除 (含参数变量判定)"""
        if Analyzer._should_exclude_item(item_name):
            return True
        if "item_name" in df.columns and "param_variable" in df.columns:
            sub = df[df["item_name"] == item_name]
            if not sub.empty:
                pv = sub.iloc[0].get("param_variable")
                if pv is not None:
                    return Analyzer._should_exclude_param(pv)
        return False

    @staticmethod
    def _extract_sn(seq_val) -> str:
        """从序列号字符串中提取后7位数字作为 SN

        规则: 取序列号字符串中所有数字字符, 取末尾7位。
        不足7位时返回全部数字; 无数字时返回空串。
        例: "W11-25080019" -> 数字 "1125080019" -> 后7位 "5080019"
        """
        if seq_val is None:
            return ""
        digits = re.findall(r"\d", str(seq_val))
        if not digits:
            return ""
        if len(digits) <= 7:
            return "".join(digits)
        return "".join(digits[-7:])

    @staticmethod
    def _parse_datetime(v):
        """将单元格值解析为 datetime 对象, 失败返回 None

        支持的类型: datetime / pandas.Timestamp / 字符串 (多种格式)
        """
        if v is None:
            return None
        if isinstance(v, float) and math.isnan(v):
            return None
        if isinstance(v, datetime):
            return v
        if isinstance(v, pd.Timestamp):
            return v.to_pydatetime()
        s = str(v).strip()
        if not s or s.lower() in ("nan", "none", "nat"):
            return None
        # 先尝试常见显式格式
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S",
                    "%Y-%m-%d %H:%M", "%Y/%m/%d %H:%M",
                    "%Y-%m-%d %H:%M", "%Y%m%d %H:%M"):
            try:
                return datetime.strptime(s, fmt)
            except ValueError:
                continue
        # 退回 pandas 解析 (能处理各种变体)
        try:
            ts = pd.to_datetime(s, errors="coerce")
            if pd.isna(ts):
                return None
            # pandas 解析"13:20"这类纯时间会得到 1900-01-01, 视为无日期
            dt = ts.to_pydatetime()
            if dt.year == 1900 and dt.month == 1 and dt.day == 1:
                return None
            return dt
        except (ValueError, TypeError):
            return None

    @staticmethod
    def _has_date_component(v) -> bool:
        """判断值是否含"年月日"日期部分"""
        if v is None:
            return False
        if isinstance(v, float) and math.isnan(v):
            return False
        if isinstance(v, (datetime, pd.Timestamp)):
            return True
        s = str(v).strip()
        if not s:
            return False
        # 含"YYYY-MM-DD"或"YYYY/MM/DD"或"YYYYMMDD"或"YYYY年MM月DD日"等模式
        if re.search(r"\d{4}[/\-年]\d{1,2}[/\-月]\d{1,2}", s):
            return True
        if re.search(r"\d{8}", s):  # 8 位连续数字视为 YYYYMMDD
            return True
        # pandas 解析后非 1900-01-01 即认为有日期
        dt = Analyzer._parse_datetime(v)
        return dt is not None

    @staticmethod
    def _format_time_value(main_val, alt_vals) -> str:
        """格式化时间值为 "月.日-时:分" (不显示年)

        规则:
          - 若 main_val 含日期, 直接用 main_val 格式化
          - 若 main_val 不含日期 (只是时间如 "13:20"或空), 用 alt_vals 中第一个含日期的值替代
          - 若所有值都不含日期, 返回原 main_val 字符串
        """
        # 主值含日期
        if Analyzer._has_date_component(main_val):
            dt = Analyzer._parse_datetime(main_val)
            if dt is not None:
                return dt.strftime("%m.%d-%H:%M")
            return str(main_val)

        # 主值不含日期, 尝试用备选值 (结束时间/开始时间) 替代
        for alt in alt_vals:
            if Analyzer._has_date_component(alt):
                dt = Analyzer._parse_datetime(alt)
                if dt is not None:
                    return dt.strftime("%m.%d-%H:%M")

        # 全部不含日期, 返回原值字符串
        if main_val is None:
            return ""
        s = str(main_val).strip()
        if s.lower() in ("nan", "none", "nat"):
            return ""
        return s

    @staticmethod
    def get_test_items(df: pd.DataFrame) -> list:
        """获取数据中所有数值型测试项目列表 (按首次出现顺序)"""
        if "item_name" not in df.columns or df.empty:
            return []
        # 保持出现顺序
        items = df["item_name"].dropna().astype(str)
        return list(dict.fromkeys(items.tolist()))

    @staticmethod
    def _parse_spec_value(raw) -> Optional[float]:
        """将规格值解析为 float, 保留原始字符串精度以避免浮点误差

        例: "1.515" -> 1.515 (而非 1.5149999999999998)
        """
        if raw is None:
            return None
        if isinstance(raw, float) and math.isnan(raw):
            return None
        s = str(raw).strip()
        if not s or s.lower() in ("nan", "none", "nat"):
            return None
        try:
            f = float(s)
        except (ValueError, TypeError):
            return None
        if "." in s:
            decimals = len(s.split(".")[-1])
            return round(f, decimals)
        return f

    @staticmethod
    def get_item_metadata(df: pd.DataFrame, item_name: str) -> dict:
        """提取测试项目的元数据 (备注/参数变量/规格上下限)"""
        sub = df[df["item_name"] == item_name] if "item_name" in df.columns else pd.DataFrame()
        if sub.empty:
            return {}
        row = sub.iloc[0]
        su_raw = row.get("spec_upper")
        sl_raw = row.get("spec_lower")
        return {
            "item_name": item_name,
            "item_note": str(row.get("item_note", "") or ""),
            "param_variable": str(row.get("param_variable", "") or ""),
            "spec_upper": Analyzer._parse_spec_value(su_raw),
            "spec_lower": Analyzer._parse_spec_value(sl_raw),
            "spec_upper_str": str(su_raw).strip() if pd.notna(su_raw) else None,
            "spec_lower_str": str(sl_raw).strip() if pd.notna(sl_raw) else None,
        }

    @staticmethod
    def filter_by_item(df: pd.DataFrame, test_item: Optional[str] = None) -> pd.DataFrame:
        """按测试项目筛选数据"""
        if test_item and "item_name" in df.columns:
            return df[df["item_name"] == test_item].reset_index(drop=True)
        return df

    @staticmethod
    def statistics(df: pd.DataFrame) -> dict:
        """计算 item_value 列的统计指标, 含 CP/CPK"""
        if df.empty or "item_value" not in df.columns:
            return {}
        values = pd.to_numeric(df["item_value"], errors="coerce").dropna()
        if values.empty:
            return {}

        n = int(values.count())
        mean_val = float(values.mean())
        std_val = float(values.std()) if n > 1 else 0.0

        # 合格率 + CP/CPK: 若存在规格上下限
        pass_rate = None
        cp = None
        cpk = None
        if "spec_upper" in df.columns and "spec_lower" in df.columns:
            su = pd.to_numeric(df["spec_upper"], errors="coerce").dropna()
            sl = pd.to_numeric(df["spec_lower"], errors="coerce").dropna()
            if not su.empty and not sl.empty:
                upper = Analyzer._parse_spec_value(su.iloc[0])
                lower = Analyzer._parse_spec_value(sl.iloc[0])
                in_range = ((values >= lower) & (values <= upper)).sum()
                pass_rate = round(in_range / n * 100, 2)
                # CP: 过程能力指数 (双边规格)
                if std_val > 0 and upper > lower:
                    cp = (upper - lower) / (6 * std_val)
                    # CPK: 考虑均值偏移的过程能力指数
                    cpu = (upper - mean_val) / (3 * std_val) if std_val > 0 else None
                    cpl = (mean_val - lower) / (3 * std_val) if std_val > 0 else None
                    if cpu is not None and cpl is not None:
                        cpk = min(cpu, cpl)
                    elif cpu is not None:
                        cpk = cpu
                    elif cpl is not None:
                        cpk = cpl

        stats = {
            "count": n,
            "mean": mean_val,
            "std": std_val,
            "min": float(values.min()),
            "max": float(values.max()),
            "median": float(values.median()),
            "q1": float(values.quantile(0.25)),
            "q3": float(values.quantile(0.75)),
            "pass_rate": pass_rate,
            "cp": cp,
            "cpk": cpk,
        }
        return stats

    @staticmethod
    def get_stats_summary(df: pd.DataFrame, only_items=None, charts=None) -> list:
        """获取所有项目的统计汇总列表

        Args:
            df: 解析后的 DataFrame
            only_items: 仅包含这些项目名 (来自首页图表), None 则全部
            charts: 首页图表列表, 用于提取同组 series 信息

        返回: [
            {
                "item_name": "项目名称",
                "item_note": "项目备注",
                "spec_lower": 下限,
                "spec_upper": 上限,
                "mean": 平均值,
                "std": 标准差,
                "cp": CP,
                "cpk": CPK,
                "count": 数据点数,
                "pass_rate": 合格率,
                "chart_idx": 所属图表索引,
                "group_series": 同组全部 series 数据,
            },
            ...
        ]
        """
        if df.empty or "item_name" not in df.columns:
            return []

        items = Analyzer.get_test_items(df)
        # 按首页图表过滤: 仅保留图表中出现的项目
        if only_items:
            item_set = set(only_items)
            items = [it for it in items if it in item_set]

        # 构建 item_name -> (chart_idx, group_series) 映射
        item_to_chart = {}
        if charts:
            for c_idx, chart in enumerate(charts):
                series_list = chart.get("series", [])
                for s in series_list:
                    name = s.get("name", "")
                    if name:
                        item_to_chart[name] = {
                            "chart_idx": c_idx,
                            "group_series": series_list,
                            "x_axis": chart.get("x_axis", []),
                            "sn_list": chart.get("sn_list", []),
                            "spec_upper": chart.get("spec_upper"),
                            "spec_lower": chart.get("spec_lower"),
                            "spec_upper_str": chart.get("spec_upper_str"),
                            "spec_lower_str": chart.get("spec_lower_str"),
                            "chart_name": chart.get("item", ""),
                        }

        result = []
        for item in items:
            if Analyzer._should_exclude_item(item):
                continue
            sub = Analyzer.filter_by_item(df, item)
            meta = Analyzer.get_item_metadata(df, item)
            stats = Analyzer.statistics(sub)

            entry = {
                "item_name": item,
                "item_note": meta.get("item_note", ""),
                "spec_lower": meta.get("spec_lower"),
                "spec_upper": meta.get("spec_upper"),
                "spec_lower_str": meta.get("spec_lower_str"),
                "spec_upper_str": meta.get("spec_upper_str"),
                "mean": stats.get("mean"),
                "std": stats.get("std"),
                "cp": stats.get("cp"),
                "cpk": stats.get("cpk"),
                "count": stats.get("count"),
                "pass_rate": stats.get("pass_rate"),
            }
            # 附加分组信息
            if item in item_to_chart:
                info = item_to_chart[item]
                entry["chart_idx"] = info["chart_idx"]
                entry["chart_name"] = info["chart_name"]
                entry["group_series"] = info["group_series"]
                entry["group_x_axis"] = info["x_axis"]
                entry["group_sn_list"] = info["sn_list"]
                entry["group_spec_upper"] = info["spec_upper"]
                entry["group_spec_lower"] = info["spec_lower"]
                entry["group_spec_upper_str"] = info.get("spec_upper_str")
                entry["group_spec_lower_str"] = info.get("spec_lower_str")
            result.append(entry)
        return result

    @staticmethod
    def chart_data(df: pd.DataFrame, test_item: Optional[str] = None) -> dict:
        """生成 ECharts 所需的曲线图数据

        返回: {
            item, meta, x_axis, sn_list, values, stats,
            spec_upper, spec_lower
        }
        - sn_list: 每个数据点对应的 SN (序列号后7位数字), 用于 tooltip 显示
        """
        filtered = Analyzer.filter_by_item(df, test_item)

        # X 轴优先用时间 (fix_time), 若不含日期则用 fix_end_time 或 fix_start_time 替代
        # 格式化为 "月.日-时:分" (不显示年)
        has_time = ("fix_time" in filtered.columns
                    or "fix_end_time" in filtered.columns
                    or "fix_start_time" in filtered.columns)
        if has_time:
            main_vals = filtered["fix_time"].tolist() if "fix_time" in filtered.columns else [None] * len(filtered)
            end_vals = filtered["fix_end_time"].tolist() if "fix_end_time" in filtered.columns else []
            start_vals = filtered["fix_start_time"].tolist() if "fix_start_time" in filtered.columns else []
            x_axis = []
            for i in range(len(filtered)):
                main = main_vals[i] if i < len(main_vals) else None
                alts = []
                if i < len(end_vals):
                    alts.append(end_vals[i])
                if i < len(start_vals):
                    alts.append(start_vals[i])
                x_axis.append(Analyzer._format_time_value(main, alts))
        elif "sequence" in filtered.columns:
            x_axis = filtered["sequence"].astype(str).tolist()
        else:
            x_axis = [str(i + 1) for i in range(len(filtered))]

        # SN: 始终从 sequence 提取后7位数字, 与 X 轴选择无关
        if "sequence" in filtered.columns:
            sn_list = [Analyzer._extract_sn(s) for s in filtered["sequence"].tolist()]
        else:
            sn_list = [""] * len(filtered)

        values = pd.to_numeric(filtered["item_value"], errors="coerce").tolist()
        stats = Analyzer.statistics(filtered)
        meta = Analyzer.get_item_metadata(df, test_item) if test_item else {}

        return {
            "item": test_item or "全部数据",
            "meta": meta,
            "x_axis": x_axis,
            "sn_list": sn_list,
            "values": values,
            "stats": stats,
            "spec_upper": meta.get("spec_upper"),
            "spec_lower": meta.get("spec_lower"),
            "spec_upper_str": meta.get("spec_upper_str"),
            "spec_lower_str": meta.get("spec_lower_str"),
        }

    @staticmethod
    def grouped_chart_data(df: pd.DataFrame, merge_items: list) -> list:
        """按合并组生成图表数据列表

        merge_items 格式: [
            {"name": "组名", "items": ["项A", "项B", ...]},
            ...
        ]
        - 在 merge_items 中出现的项, 与同组其他项绘制到同一图表(多 series)
        - 不在任何合并组中的项, 单独作为一个图表(单 series)
        - 返回的图表对象新增 series 字段: [{name, values, stats}, ...]
        """
        if not merge_items:
            return Analyzer.all_items_chart_data(df)

        # 收集所有被合并的项名, 用于后续判断
        merged_set = set()
        for g in merge_items:
            for it in g.get("items", []):
                merged_set.add(it)

        all_items = Analyzer.get_test_items(df)
        charts = []
        processed = set()

        # 1. 先处理合并组
        for group in merge_items:
            group_name = group.get("name", "")
            group_items = [it for it in group.get("items", [])
                           if it in all_items and not Analyzer._should_exclude_item(it)]
            if not group_items:
                continue

            # 以第一个项的 X 轴为基准 (序列号)
            base_chart = Analyzer.chart_data(df, group_items[0])
            base_meta = base_chart.get("meta") or {}
            series_list = [{
                "name": group_items[0],
                "values": base_chart["values"],
                "stats": base_chart["stats"],
                "item_note": base_meta.get("item_note", ""),
            }]
            # 其他项按相同序列号对齐 (若序列号一致则直接用, 否则用自己的 x_axis)
            for extra in group_items[1:]:
                ec = Analyzer.chart_data(df, extra)
                ec_meta = ec.get("meta") or {}
                series_list.append({
                    "name": extra,
                    "values": ec["values"],
                    "stats": ec["stats"],
                    "item_note": ec_meta.get("item_note", ""),
                })

            # 合并规格上下限: 取组内任一非空值 (按设计应一致)
            spec_upper = None
            spec_lower = None
            spec_upper_str = None
            spec_lower_str = None
            for it in group_items:
                m = Analyzer.get_item_metadata(df, it)
                if m.get("spec_upper") is not None:
                    spec_upper = m["spec_upper"]
                    spec_upper_str = m.get("spec_upper_str")
                if m.get("spec_lower") is not None:
                    spec_lower = m["spec_lower"]
                    spec_lower_str = m.get("spec_lower_str")

            # 合并 meta: 取组内第一个项的元数据作为代表
            meta = Analyzer.get_item_metadata(df, group_items[0])

            charts.append({
                "item": group_name or " / ".join(group_items),
                "meta": meta,
                "x_axis": base_chart["x_axis"],
                "sn_list": base_chart.get("sn_list", []),  # 透传 SN 列表
                "values": base_chart["values"],  # 兼容旧字段(第一个 series)
                "stats": base_chart["stats"],    # 兼容旧字段
                "spec_upper": spec_upper,
                "spec_lower": spec_lower,
                "spec_upper_str": spec_upper_str,
                "spec_lower_str": spec_lower_str,
                "series": series_list,           # 新字段: 多条曲线
            })
            for it in group_items:
                processed.add(it)

        # 2. 处理未合并的项 (单独图表)
        for item in all_items:
            if item in processed:
                continue
            if Analyzer._should_exclude_item(item):
                continue
            cd = Analyzer.chart_data(df, item)
            cd_meta = cd.get("meta") or {}
            cd["series"] = [{
                "name": item,
                "values": cd["values"],
                "stats": cd["stats"],
                "item_note": cd_meta.get("item_note", ""),
            }]
            charts.append(cd)

        return charts

    @staticmethod
    def all_items_chart_data(df: pd.DataFrame) -> list:
        """为每个数值型测试项目生成图表数据列表"""
        items = [it for it in Analyzer.get_test_items(df)
                 if not Analyzer._should_exclude_item(it)]
        if not items:
            return [Analyzer.chart_data(df, None)]
        charts = []
        for item in items:
            cd = Analyzer.chart_data(df, item)
            cd_meta = cd.get("meta") or {}
            cd["series"] = [{
                "name": item,
                "values": cd["values"],
                "stats": cd["stats"],
                "item_note": cd_meta.get("item_note", ""),
            }]
            charts.append(cd)
        return charts

    @staticmethod
    def auto_grouped_chart_data(df: pd.DataFrame) -> list:
        """按数据提取规则3+4 自动分组, 同组测试项目绘制到同一图表

        规则3: 仅保留有上下限 (至少其一不为 None) 的数值型测试项目作为分析数据
        规则4: 按 (spec_lower, spec_upper) 元组分组, 元组相同的项目合并为同一图表
               (None 与 None 视为相同, 即所有"仅有上限"或"仅有下限"的项目分别合并)

        - 分组 key: (spec_lower, spec_upper); None 保留为 None 参与分组
        - 组名: 规格区间 "下限~上限 (N项)" 或 单边规格描述
        - 每组输出一个图表, 内含 N 条 series (N = 组内项目数)
        - 组排序: 按组内首个项目的首次出现顺序 (保持 Excel 列顺序)
        """
        all_items = Analyzer.get_test_items(df)
        if not all_items:
            return [Analyzer.chart_data(df, None)]

        # 1. 收集每个测试项目的规格上下限, 并按规则3过滤 (无任何规格限的项目跳过)
        groups: dict = {}       # key=(lower, upper) -> [item_name, ...]
        group_meta: dict = {}    # key -> 代表 meta
        group_first_idx: dict = {}  # key -> 组内首项在 all_items 中的索引 (用于排序)
        for idx, it in enumerate(all_items):
            # 跳过不参与分析的项目 (如 CAN_报文定义/数字量状态量)
            if Analyzer._item_excluded(df, it):
                continue
            m = Analyzer.get_item_metadata(df, it)
            su_raw = m.get("spec_upper")
            sl_raw = m.get("spec_lower")
            # 规则3: 至少有 spec_lower 或 spec_upper 其一 (都不为 None 的项才入选)
            if su_raw is None and sl_raw is None:
                continue
            # 元组 key, 保留 None (None 与 None 视为相同, 可合并)
            key = (sl_raw, su_raw)
            groups.setdefault(key, []).append(it)
            if key not in group_meta:
                group_meta[key] = m
                group_first_idx[key] = idx

        # 2. 为每组生成图表, 按组内首项首次出现顺序排序 (规则4: 保持 Excel 列顺序)
        charts = []
        for key in sorted(groups.keys(), key=lambda k: group_first_idx[k]):
            items_in_group = groups[key]
            if not items_in_group:
                continue

            base_chart = Analyzer.chart_data(df, items_in_group[0])
            base_meta = base_chart.get("meta") or {}
            series_list = [{
                "name": items_in_group[0],
                "values": base_chart["values"],
                "stats": base_chart["stats"],
                "item_note": base_meta.get("item_note", ""),
            }]
            for extra in items_in_group[1:]:
                ec = Analyzer.chart_data(df, extra)
                ec_meta = ec.get("meta") or {}
                series_list.append({
                    "name": extra,
                    "values": ec["values"],
                    "stats": ec["stats"],
                    "item_note": ec_meta.get("item_note", ""),
                })

            meta = group_meta[key]
            su = meta.get("spec_upper")
            sl = meta.get("spec_lower")
            # 组名: 显示组内所有项目名称 (项目名称来自 Excel 表头 item_name 行)
            # 规格信息保留在 spec_upper/spec_lower 字段, 由前端徽章显示
            group_name = " / ".join(items_in_group)

            charts.append({
                "item": group_name,
                "meta": meta,
                "x_axis": base_chart["x_axis"],
                "sn_list": base_chart.get("sn_list", []),  # 透传 SN 列表
                "values": base_chart["values"],
                "stats": base_chart["stats"],
                "spec_upper": su,
                "spec_lower": sl,
                "series": series_list,
            })

        return charts
