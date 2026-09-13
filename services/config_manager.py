"""产品类别配置管理服务

以 JSON 文件形式保存各产品类别的解析配置，便于后续同款产品直接调用。
"""
import os
import json
from typing import Optional


class ConfigManager:
    """管理 configs/ 目录下的产品类别 JSON 配置"""

    def __init__(self, config_folder: str):
        self.folder = config_folder
        os.makedirs(self.folder, exist_ok=True)

    def _path(self, name: str) -> str:
        """根据产品名称返回配置文件路径"""
        safe = name.replace("/", "_").replace("\\", "_").strip()
        return os.path.join(self.folder, f"{safe}.json")

    def list_configs(self) -> list:
        """列出所有已保存的产品配置"""
        configs = []
        for fname in os.listdir(self.folder):
            if fname.endswith(".json"):
                path = os.path.join(self.folder, fname)
                try:
                    with open(path, "r", encoding="utf-8") as f:
                        data = json.load(f)
                    configs.append({
                        "name": data.get("category_name", fname[:-5]),
                        "filename": fname,
                        "test_items": data.get("test_items", []),
                    })
                except (json.JSONDecodeError, IOError):
                    continue
        return configs

    def load(self, name: str) -> Optional[dict]:
        """读取指定产品类别的配置"""
        path = self._path(name)
        if not os.path.exists(path):
            return None
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)

    def save(self, config: dict) -> str:
        """保存产品类别配置，返回保存的文件名"""
        name = config.get("category_name", "未命名产品")
        path = self._path(name)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        return os.path.basename(path)

    def delete(self, name: str) -> bool:
        """删除指定产品类别配置"""
        path = self._path(name)
        if os.path.exists(path):
            os.remove(path)
            return True
        return False

    @staticmethod
    def default_config() -> dict:
        """返回一个默认配置模板，供前端编辑

        Excel 格式 (依据 测试分析001.md excel数据表格式说明):

        - skip_top_rows: 起始跳过行 (例如 SC 板第0行是 "Program Name:" 标题)
        - 接下来的 header_rows 行: 测试项目元数据表头
            行0: 项目名称 (若某项目占多列且同名, 会自动组合参数变量成唯一项名)
            行1: 项目备注
            行2: 参数变量 (也可作为子项，如 CH1电压/CH1电流)
            行3: 规格上限
            行4: 规格下限   ('*' 或 '-' 表示无限)
        - fixed_columns_header_row: 相对 skip_top_rows 后的行号, 若 >=0 则从此行读取固定列名
            (例如真实 SC 板测试表, 固定列 "序列号/工单号/名称/测试员/工装编号/结论/时间" 在元数据之后)
        - 固定列与测试值列之后: 产品数据 (每个产品一行, 序列号为空则跳过)
        - 文本型测试值自动跳过, 仅解析数值型
        """
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
