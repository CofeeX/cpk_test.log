"""CPK-PCBA - Flask 主应用

路由:
  GET  /                    首页
  POST /api/upload          上传并解析 Excel
  GET  /api/chart-data      获取图表数据(可按测试项目筛选)
  GET  /api/test-items      获取测试项目列表
  GET  /api/configs          列出所有产品配置
  GET  /api/config/<name>    获取指定产品配置
  POST /api/config           保存产品配置
  DELETE /api/config/<name>  删除产品配置
  GET  /api/sheets           获取已上传文件的工作表列表
  GET  /api/history          获取上传历史记录
"""
import os
import pickle
import uuid
import json
import threading
import webbrowser

import pandas as pd
from flask import (
    Flask, render_template, request, jsonify, session, send_from_directory
)

from config import Config
from models import db, UploadRecord
from services.excel_parser import ExcelParser
from services.analyzer import Analyzer
from services.config_manager import ConfigManager

app = Flask(__name__)
app.config.from_object(Config)

# 确保目录存在
os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
os.makedirs(app.config["CONFIG_FOLDER"], exist_ok=True)

db.init_app(app)

config_manager = ConfigManager(app.config["CONFIG_FOLDER"])


# --------------------------------------------------------------------------- #
#  页面路由
# --------------------------------------------------------------------------- #
@app.route("/")
def index():
    """首页：文件选择、分析按钮、测试项目曲线图"""
    has_data = bool(session.get("df_file"))
    filename = session.get("filename", "")
    product_category = session.get("product_category", "")
    return render_template("index.html",
                           configs=config_manager.list_configs(),
                           has_data=has_data,
                           filename=filename,
                           product_category=product_category)


@app.route("/stats")
def stats_page():
    """数据统计页：列表展示所有项目的统计汇总"""
    has_data = bool(session.get("df_file"))
    filename = session.get("filename", "")
    product_category = session.get("product_category", "")
    return render_template("stats.html",
                           has_data=has_data,
                           filename=filename,
                           product_category=product_category)


# --------------------------------------------------------------------------- #
#  辅助函数
# --------------------------------------------------------------------------- #
def _allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in Config.ALLOWED_EXTENSIONS


def _save_df(df: pd.DataFrame) -> str:
    """将 DataFrame 序列化到临时文件，返回文件名"""
    fname = f"{uuid.uuid4().hex}.pkl"
    path = os.path.join(Config.UPLOAD_FOLDER, fname)
    with open(path, "wb") as f:
        pickle.dump(df, f)
    return fname


def _save_charts(charts: list) -> str:
    """将图表数据序列化到临时文件，返回文件名"""
    fname = f"{uuid.uuid4().hex}_charts.pkl"
    path = os.path.join(Config.UPLOAD_FOLDER, fname)
    with open(path, "wb") as f:
        pickle.dump(charts, f)
    return fname


def _load_charts() -> list:
    """从 session 中记录的文件加载图表数据"""
    pkl_name = session.get("charts_file")
    if not pkl_name:
        return []
    path = os.path.join(Config.UPLOAD_FOLDER, pkl_name)
    if not os.path.exists(path):
        return []
    with open(path, "rb") as f:
        return pickle.load(f)


def _load_df() -> pd.DataFrame:
    """从 session 中记录的文件加载 DataFrame"""
    pkl_name = session.get("df_file")
    if not pkl_name:
        return pd.DataFrame()
    path = os.path.join(Config.UPLOAD_FOLDER, pkl_name)
    if not os.path.exists(path):
        return pd.DataFrame()
    with open(path, "rb") as f:
        return pickle.load(f)


# --------------------------------------------------------------------------- #
#  上传与解析
# --------------------------------------------------------------------------- #
@app.route("/api/upload", methods=["POST"])
def upload():
    """上传 Excel 文件，按选定的产品配置解析并返回测试项目列表"""
    if "file" not in request.files:
        return jsonify({"error": "未选择文件"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "未选择文件"}), 400

    if not _allowed_file(file.filename):
        return jsonify({"error": "仅支持 .xlsx / .xls 文件"}), 400

    # 保存原始文件
    filename = file.filename
    saved_path = os.path.join(Config.UPLOAD_FOLDER, filename)
    file.save(saved_path)

    # 获取选定的产品配置（功能3: 根据不同产品类别不同解析入口）
    product_name = request.form.get("product_category", "").strip()
    product_config = config_manager.load(product_name) if product_name else None

    if product_config:
        parser = ExcelParser(product_config)
    else:
        # 无配置时从导入文件自动检测配置
        product_config = ExcelParser.auto_detect_config(saved_path)
        parser = ExcelParser(product_config)

    try:
        df = parser.parse(saved_path)
    except Exception as e:
        return jsonify({"error": f"解析失败: {str(e)}"}), 400

    # 提取公共信息 (Program Name 等顶部标题 + 工单号/测试员/工装编号等)
    common_info = {}
    if hasattr(parser, "program_info") and parser.program_info:
        common_info.update(parser.program_info)
    # 从数据行补充工单号/测试员/工装编号 (取首条非空)
    for src_col, key in [
        ("fix_work_order", "工单号"),
        ("fix_tester", "测试员"),
        ("fix_fixture", "工装编号"),
    ]:
        if src_col in df.columns and key not in common_info:
            vals = df[src_col].dropna().astype(str)
            vals = vals[vals.str.strip() != ""]
            if not vals.empty:
                common_info[key] = str(vals.iloc[0])

    # 序列化 DataFrame 到临时文件
    pkl_name = _save_df(df)
    session["df_file"] = pkl_name
    session["filename"] = filename
    product_label = product_name or "自动检测配置"
    session["product_category"] = product_label

    # 记录到数据库
    test_items = Analyzer.get_test_items(df)
    record = UploadRecord(
        filename=filename,
        saved_path=saved_path,
        product_category=product_label,
        total_rows=len(df),
        test_items=json.dumps(test_items, ensure_ascii=False),
    )
    db.session.add(record)
    db.session.commit()

    # 返回测试项目列表及初始图表数据（功能4 & 功能5）
    # 图表合并优先级:
    #   1) 配置显式声明 merge_items -> 按配置分组
    #   2) test_items 为空 (白名单全开) -> 按 (规格下限, 规格上限) 自动分组
    #   3) 否则 -> 每项一张图
    merge_items = product_config.get("merge_items", []) if product_config else []
    test_items_cfg = product_config.get("test_items", []) if product_config else []
    if merge_items:
        charts = Analyzer.grouped_chart_data(df, merge_items)
    elif not test_items_cfg:
        charts = Analyzer.auto_grouped_chart_data(df)
    else:
        charts = Analyzer.all_items_chart_data(df)

    # 调试限制已移除: 显示全部图表 (按规则3+4 自动分组后的所有图表)
    # _DEBUG_MAX_CHARTS = 4
    # if _DEBUG_MAX_CHARTS and len(charts) > _DEBUG_MAX_CHARTS:
    #     charts = charts[:_DEBUG_MAX_CHARTS]

    # 统计: 产品数量(序列号去重)、分析项目数、图表数
    seq_col = "fix_sequence" if "fix_sequence" in df.columns else "sequence"
    product_count = int(pd.Series(df[seq_col].unique()).dropna().shape[0]) if seq_col in df.columns else len(df)
    # 规则3: 分析项目数 = 有上下限 (至少其一) 的数值型测试项目数
    # (与 auto_grouped_chart_data 内部过滤逻辑一致: 仅 spec_lower 和 spec_upper 同时为 None 时排除)
    analysis_item_count = 0
    if "item_name" in df.columns and "spec_lower" in df.columns and "spec_upper" in df.columns:
        # 按 item_name 聚合, 取每组第一条判断规格限
        for _item, _grp in df.groupby("item_name", sort=False):
            # 数字量/状态量等不参与分析的项目不计数
            if Analyzer._item_excluded(df, _item):
                continue
            _first = _grp.iloc[0]
            _sl = _first.get("spec_lower")
            _su = _first.get("spec_upper")
            try:
                _sl_v = float(_sl) if pd.notna(_sl) and str(_sl).strip() else None
            except (TypeError, ValueError):
                _sl_v = None
            try:
                _su_v = float(_su) if pd.notna(_su) and str(_su).strip() else None
            except (TypeError, ValueError):
                _su_v = None
            if _sl_v is not None or _su_v is not None:
                analysis_item_count += 1
    else:
        analysis_item_count = len(test_items)
    chart_count = len(charts)

    # 提取图表中实际显示的项目名, 存入 session 供统计页过滤
    chart_item_names = set()
    for chart in charts:
        for s in chart.get("series", []):
            name = s.get("name", "")
            if name:
                chart_item_names.add(name)
    session["chart_item_names"] = sorted(chart_item_names)

    # 保存图表数据到 session, 供首页恢复
    charts_file = _save_charts(charts)
    session["charts_file"] = charts_file

    return jsonify({
        "filename": filename,
        "product_category": product_label,
        "total_rows": len(df),
        "product_count": product_count,
        "analysis_item_count": analysis_item_count,
        "chart_count": chart_count,
        "common_info": common_info,
        "test_items": test_items,
        "charts": charts,
        "record_id": record.id,
    })


# --------------------------------------------------------------------------- #
#  图表数据
# --------------------------------------------------------------------------- #
@app.route("/api/chart-data", methods=["GET"])
def chart_data():
    """获取指定测试项目的曲线图与统计数据（功能5: 筛选目标项目）"""
    test_item = request.args.get("item")
    df = _load_df()
    if df.empty:
        return jsonify({"error": "无数据，请先上传文件"}), 400

    data = Analyzer.chart_data(df, test_item)
    return jsonify(data)


@app.route("/api/test-items", methods=["GET"])
def test_items():
    """获取当前数据的测试项目列表"""
    df = _load_df()
    if df.empty:
        return jsonify({"test_items": []})
    items = Analyzer.get_test_items(df)
    stats = {item: Analyzer.statistics(Analyzer.filter_by_item(df, item)) for item in items}
    return jsonify({"test_items": items, "stats": stats})


@app.route("/api/stats", methods=["GET"])
def get_stats():
    """获取首页图表中实际显示项目的统计汇总列表"""
    df = _load_df()
    if df.empty:
        return jsonify({"stats": [], "total_count": 0})
    # 仅展示首页图表中出现的项目
    chart_item_names = session.get("chart_item_names", [])
    charts = _load_charts()
    stats = Analyzer.get_stats_summary(df, only_items=chart_item_names, charts=charts)
    return jsonify({"stats": stats, "total_count": len(stats)})


@app.route("/api/session-charts", methods=["GET"])
def session_charts():
    """获取 session 中保存的图表数据, 供首页恢复"""
    charts = _load_charts()
    return jsonify({
        "charts": charts,
        "has_data": bool(charts),
        "filename": session.get("filename", ""),
        "product_category": session.get("product_category", ""),
    })


# --------------------------------------------------------------------------- #
#  产品配置管理（功能6: 按产品类别保存配置）
# --------------------------------------------------------------------------- #
@app.route("/api/configs", methods=["GET"])
def list_configs():
    return jsonify({"configs": config_manager.list_configs()})


@app.route("/api/config/<path:name>", methods=["GET"])
def get_config(name):
    cfg = config_manager.load(name)
    if cfg is None:
        return jsonify({"error": "配置不存在"}), 404
    return jsonify(cfg)


@app.route("/api/config", methods=["POST"])
def save_config():
    cfg = request.get_json()
    if not cfg or "category_name" not in cfg:
        return jsonify({"error": "配置需包含 category_name"}), 400
    fname = config_manager.save(cfg)
    return jsonify({"message": "配置已保存", "filename": fname})


@app.route("/api/config/<path:name>", methods=["DELETE"])
def delete_config(name):
    if config_manager.delete(name):
        return jsonify({"message": "配置已删除"})
    return jsonify({"error": "配置不存在"}), 404


@app.route("/api/config/default", methods=["GET"])
def default_config():
    """返回默认配置模板，供前端编辑"""
    return jsonify(ConfigManager.default_config())


# --------------------------------------------------------------------------- #
#  工作表与历史记录
# --------------------------------------------------------------------------- #
@app.route("/api/sheets", methods=["GET"])
def get_sheets():
    """获取当前上传文件的工作表列表，便于配置时选择"""
    filename = session.get("filename")
    if not filename:
        return jsonify({"error": "无上传文件"}), 400
    path = os.path.join(Config.UPLOAD_FOLDER, filename)
    if not os.path.exists(path):
        return jsonify({"error": "文件不存在"}), 404
    try:
        sheets = ExcelParser.get_sheet_names(path)
    except Exception as e:
        return jsonify({"error": str(e)}), 400
    return jsonify({"sheets": sheets})


@app.route("/api/history", methods=["GET"])
def history():
    records = UploadRecord.query.order_by(UploadRecord.created_at.desc()).limit(50).all()
    return jsonify({"history": [r.to_dict() for r in records]})


# --------------------------------------------------------------------------- #
#  启动
# --------------------------------------------------------------------------- #
with app.app_context():
    db.create_all()


def _open_browser():
    """延迟打开默认浏览器访问应用首页"""
    try:
        webbrowser.open("http://127.0.0.1:5000/", new=2)
    except Exception:
        pass


if __name__ == "__main__":
    import sys
    # 打包为 exe 后(sys.frozen=True) 用生产模式, 直接运行 Python 时用 debug 模式
    DEBUG = not getattr(sys, 'frozen', False)
    # debug 模式下 werkzeug reloader 会启动子进程, 仅在子进程打开浏览器避免重复
    # 打包后(无 reloader) 直接打开
    should_open = (not DEBUG) or (os.environ.get("WERKZEUG_RUN_MAIN") == "true")
    if should_open:
        threading.Timer(1.5, _open_browser).start()
    app.run(debug=DEBUG, host="0.0.0.0", port=5000)
