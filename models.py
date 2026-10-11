from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
import json

db = SQLAlchemy()


class Resume(db.Model):
    """Stores analyzed resume data"""
    __tablename__ = 'resumes'

    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(255), nullable=False)
    file_path = db.Column(db.String(500), nullable=True)
    file_size = db.Column(db.Integer, nullable=True)
    file_type = db.Column(db.String(20), nullable=True)

    # Extracted candidate info
    name = db.Column(db.String(200), nullable=True)
    email = db.Column(db.String(200), nullable=True)
    contact = db.Column(db.String(50), nullable=True)
    experience = db.Column(db.String(100), nullable=True)

    # JSON-encoded lists
    skills_json = db.Column(db.Text, nullable=True)
    role_analysis_json = db.Column(db.Text, nullable=True)
    top_role = db.Column(db.String(200), nullable=True)
    top_role_match = db.Column(db.Float, nullable=True)

    # Raw text extracted from resume
    raw_text = db.Column(db.Text, nullable=True)

    # Metadata
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)
    ip_address = db.Column(db.String(50), nullable=True)

    @property
    def skills(self):
        return json.loads(self.skills_json) if self.skills_json else []

    @skills.setter
    def skills(self, value):
        self.skills_json = json.dumps(value or [])

    @property
    def role_analysis(self):
        return json.loads(self.role_analysis_json) if self.role_analysis_json else []

    @role_analysis.setter
    def role_analysis(self, value):
        self.role_analysis_json = json.dumps(value or [])

    def to_dict(self, include_text=False):
        data = {
            'id': self.id,
            'filename': self.filename,
            'name': self.name,
            'email': self.email,
            'contact': self.contact,
            'experience': self.experience,
            'skills': self.skills,
            'top_role': self.top_role,
            'top_role_match': self.top_role_match,
            'role_analysis': self.role_analysis,
            'uploaded_at': self.uploaded_at.isoformat() if self.uploaded_at else None,
        }
        if include_text:
            data['raw_text'] = self.raw_text
        return data

    def __repr__(self):
        return f'<Resume {self.id}: {self.name or "Unknown"}>'