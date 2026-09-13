"""数据库模型"""
from datetime import datetime
from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


class UploadRecord(db.Model):
    """上传记录表"""
    __tablename__ = "upload_records"

    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(255), nullable=False)          # 原始文件名
    saved_path = db.Column(db.String(512), nullable=False)          # 服务器存储路径
    product_category = db.Column(db.String(100))                    # 产品类别
    total_rows = db.Column(db.Integer, default=0)                  # 数据总行数
    test_items = db.Column(db.Text)                                 # 测试项目(JSON 字符串)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self) -> dict:
        import json
        return {
            "id": self.id,
            "filename": self.filename,
            "product_category": self.product_category,
            "total_rows": self.total_rows,
            "test_items": json.loads(self.test_items) if self.test_items else [],
            "created_at": self.created_at.strftime("%Y-%m-%d %H:%M:%S"),
        }
