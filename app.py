from flask import Flask, request, jsonify, render_template
import os
import re
import tempfile
from PyPDF2 import PdfReader
from docx import Document

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Define job roles and their required skills
JOB_ROLES = {
    "Software Engineer": {
        "skills": ["python", "java", "c++", "data structures", "algorithms", "git", "sql", "oop", "problem solving"],
        "extra_skills": ["system design", "microservices", "cloud computing", "ci/cd", "testing", "docker"],
        "description": "Develops and maintains software applications, writes clean and efficient code."
    },
    "Web Developer": {
        "skills": ["html", "css", "javascript", "react", "angular", "vue", "node.js", "express", "bootstrap"],
        "extra_skills": ["typescript", "webpack", "graphql", "responsive design", "seo", "accessibility"],
        "description": "Builds and maintains websites and web applications with focus on user experience."
    },
    "Data Scientist": {
        "skills": ["python", "r", "sql", "machine learning", "deep learning", "nlp", "computer vision", "statistics", "pandas", "numpy"],
        "extra_skills": ["big data", "spark", "hadoop", "cloud computing", "data visualization", "tableau"],
        "description": "Analyzes complex data to extract insights, builds predictive models."
    },
    "DevOps Engineer": {
        "skills": ["aws", "azure", "gcp", "docker", "kubernetes", "jenkins", "git", "linux", "ci/cd", "ansible"],
        "extra_skills": ["terraform", "prometheus", "grafana", "elk stack", "security", "networking"],
        "description": "Manages infrastructure, automates deployment processes, ensures system reliability."
    },
    "Data Analyst": {
        "skills": ["sql", "excel", "python", "r", "tableau", "power bi", "statistics", "data visualization", "analytical"],
        "extra_skills": ["big data", "etl", "data warehousing", "business intelligence", "reporting"],
        "description": "Interprets data, creates reports and dashboards to support business decisions."
    },
    "Frontend Developer": {
        "skills": ["html", "css", "javascript", "react", "angular", "vue", "sass", "less", "bootstrap", "jquery"],
        "extra_skills": ["typescript", "webpack", "babel", "state management", "ui/ux", "responsive design"],
        "description": "Creates user interfaces, ensures website visual elements and interactions work smoothly."
    },
    "Backend Developer": {
        "skills": ["python", "java", "c#", "node.js", "express", "django", "flask", "spring", "sql", "mongodb"],
        "extra_skills": ["microservices", "graphql", "redis", "kafka", "rabbitmq", "security"],
        "description": "Builds server-side logic, APIs, and database interactions."
    },
    "Full Stack Developer": {
        "skills": ["html", "css", "javascript", "react", "node.js", "python", "sql", "mongodb", "git", "rest api"],
        "extra_skills": ["typescript", "docker", "aws", "graphql", "testing", "ci/cd"],
        "description": "Works on both frontend and backend development, creates complete web applications."
    },
    "Mobile Developer": {
        "skills": ["react native", "flutter", "swift", "kotlin", "java", "android", "ios", "mobile", "firebase"],
        "extra_skills": ["ui/ux design", "app store deployment", "performance optimization", "testing"],
        "description": "Builds mobile applications for iOS and Android platforms."
    },
    "Machine Learning Engineer": {
        "skills": ["python", "tensorflow", "pytorch", "scikit-learn", "machine learning", "deep learning", "nlp", "computer vision"],
        "extra_skills": ["mlops", "docker", "kubernetes", "big data", "cloud computing", "model deployment"],
        "description": "Designs and implements machine learning models, deploys ML solutions."
    },
    "Cloud Engineer": {
        "skills": ["aws", "azure", "gcp", "docker", "kubernetes", "linux", "python", "networking", "security"],
        "extra_skills": ["terraform", "ansible", "prometheus", "cloud architecture", "cost optimization"],
        "description": "Designs and manages cloud infrastructure, ensures scalability and security."
    },
    "Business Analyst": {
        "skills": ["excel", "sql", "python", "tableau", "business analysis", "requirements gathering", "documentation"],
        "extra_skills": ["agile", "scrum", "project management", "communication", "stakeholder management"],
        "description": "Analyzes business needs, defines requirements, and bridges gap between stakeholders and developers."
    },
    "Project Manager": {
        "skills": ["agile", "scrum", "project management", "leadership", "communication", "jira", "confluence"],
        "extra_skills": ["budgeting", "risk management", "stakeholder management", "reporting", "scheduling"],
        "description": "Leads projects, coordinates teams, ensures projects are delivered on time and within budget."
    }
}

def extract_text_from_pdf(file_path):
    """Extract text from PDF using PyPDF2"""
    text = ""
    try:
        with open(file_path, 'rb') as file:
            reader = PdfReader(file)
            for page in reader.pages:
                text += page.extract_text() or ""
    except Exception as e:
        print(f"Error reading PDF: {e}")
    return text

def extract_text_from_docx(file_path):
    """Extract text from DOCX file"""
    text = ""
    try:
        doc = Document(file_path)
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
    except Exception as e:
        print(f"Error reading DOCX: {e}")
    return text

def extract_email(text):
    """Extract email address from text"""
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    emails = re.findall(email_pattern, text)
    return emails[0] if emails else "Not found"

def extract_phone(text):
    """Extract phone number from text"""
    phone_patterns = [
        r'\+\d{1,3}[-.\s]?\(?\d{1,4}\)?[-.\s]?\d{1,4}[-.\s]?\d{1,9}',
        r'\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}',
        r'\d{3}[-.\s]\d{3}[-.\s]\d{4}',
        r'\d{10,15}'
    ]
    
    for pattern in phone_patterns:
        matches = re.findall(pattern, text)
        if matches:
            phone = matches[0].strip()
            phone = re.sub(r'\s+', ' ', phone)
            return phone
    return "Not found"

def extract_name(text):
    """Extract name using regex patterns"""
    name_patterns = [
        r'Name\s*[:|]\s*([A-Za-z\s]+)',
        r'Candidate\s*[:|]\s*([A-Za-z\s]+)',
        r'Full Name\s*[:|]\s*([A-Za-z\s]+)',
        r'^([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)'
    ]
    
    lines = text.split('\n')
    for line in lines[:20]:
        line = line.strip()
        if line and len(line) < 50 and line[0].isupper():
            words = line.split()
            if 2 <= len(words) <= 4 and all(w[0].isupper() for w in words):
                return line
    
    for pattern in name_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
    
    return "Not found"

def extract_skills(text):
    """Extract skills from text using keyword matching"""
    # Comprehensive skill list
    skill_keywords = [
        'python', 'java', 'javascript', 'react', 'angular', 'vue', 'node',
        'django', 'flask', 'spring', 'aws', 'azure', 'gcp', 'docker',
        'kubernetes', 'sql', 'mongodb', 'postgresql', 'mysql', 'redis',
        'html', 'css', 'typescript', 'jquery', 'bootstrap',
        'c++', 'c#', 'ruby', 'php', 'swift', 'kotlin', 'rust', 'go',
        'machine learning', 'deep learning', 'nlp', 'computer vision',
        'tensorflow', 'pytorch', 'scikit-learn', 'pandas', 'numpy',
        'agile', 'scrum', 'jira', 'git', 'github',
        'jenkins', 'linux', 'unix',
        'project management', 'leadership', 'communication',
        'problem solving', 'analytical', 'teamwork',
        'excel', 'powerpoint', 'word', 'outlook',
        'salesforce', 'sap', 'oracle', 'tableau', 'power bi',
        'html5', 'css3', 'sass', 'less', 'webpack', 'babel',
        'rest api', 'graphql', 'microservices', 'devops',
        'c', 'rust', 'go', 'scala', 'perl', 'matlab',
        'hadoop', 'spark', 'kafka', 'rabbitmq', 'nginx',
        'react native', 'flutter', 'kotlin', 'swift', 'mobile',
        'ui/ux', 'figma', 'photoshop', 'illustrator',
        'data analysis', 'statistics', 'data science',
        'etl', 'data warehousing', 'big data',
        'cloud', 'networking', 'security', 'cryptography'
    ]
    
    found_skills = []
    text_lower = text.lower()
    
    # Convert skills to lowercase for matching
    for skill in skill_keywords:
        # Check if skill appears as a whole word
        if re.search(r'\b' + re.escape(skill) + r'\b', text_lower):
            # Capitalize properly
            skill_title = ' '.join(word.capitalize() for word in skill.split())
            found_skills.append(skill_title)
    
    # Remove duplicates while preserving order
    unique_skills = []
    for skill in found_skills:
        if skill not in unique_skills:
            unique_skills.append(skill)
    
    return unique_skills if unique_skills else []

def extract_experience(text):
    """Extract experience information"""
    experience_patterns = [
        r'(\d+)\s*(?:years?|yrs?)\s*(?:of\s*)?experience',
        r'experience\s*(?:of\s*)?(\d+)\s*(?:years?|yrs?)',
        r'(\d+)\s*\+\s*(?:years?|yrs?)\s*(?:of\s*)?experience',
        r'(\d+)\s*(?:years?|yrs?)\s*experience',
        r'(\d+)\s*(?:years?|yrs?)\s*of\s*experience'
    ]
    
    for pattern in experience_patterns:
        matches = re.findall(pattern, text.lower())
        if matches:
            years = matches[0]
            return f"{years} years"
    
    # Look for date ranges
    date_pattern = r'(19|20)\d{2}'
    years_found = re.findall(date_pattern, text)
    if years_found:
        years = sorted(set(years_found))
        if len(years) >= 2:
            try:
                start = int(years[0])
                end = int(years[-1])
                if end > start:
                    return f"{end - start} years"
            except:
                pass
    
    return "Not specified"

def analyze_job_role(skills):
    """Analyze which job role fits best based on skills"""
    if not skills:
        return []
    
    # Convert skills to lowercase for matching
    skills_lower = [s.lower() for s in skills]
    
    # Calculate match score for each role
    role_scores = []
    
    for role_name, role_data in JOB_ROLES.items():
        required_skills = [s.lower() for s in role_data['skills']]
        
        # Find matching skills
        matched_skills = [s for s in skills_lower if any(req in s or s in req for req in required_skills)]
        
        # Calculate match percentage
        if required_skills:
            match_count = len(matched_skills)
            total_required = len(required_skills)
            match_percentage = (match_count / total_required) * 100
            
            # Find missing required skills
            missing_skills = []
            for req in required_skills:
                if not any(req in s or s in req for s in skills_lower):
                    missing_skills.append(req.title())
            
            # Find extra skills that would be valuable
            extra_skills = role_data.get('extra_skills', [])
            recommended_skills = [s.title() for s in extra_skills if not any(s in skill_lower or skill_lower in s for skill_lower in skills_lower)]
            
            role_scores.append({
                'role': role_name,
                'match_percentage': round(match_percentage, 1),
                'matched_skills': [s.title() for s in matched_skills],
                'missing_skills': missing_skills,
                'recommended_skills': recommended_skills[:5],  # Top 5 recommendations
                'description': role_data['description'],
                'total_required': total_required
            })
    
    # Sort by match percentage (highest first)
    role_scores.sort(key=lambda x: x['match_percentage'], reverse=True)
    
    return role_scores[:5]  # Return top 5 roles

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analyze', methods=['POST'])
def analyze_resume():
    try:
        if 'resume' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['resume']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        filename = file.filename.lower()
        if not (filename.endswith('.pdf') or filename.endswith('.docx')):
            return jsonify({'error': 'Only PDF and DOCX files are supported'}), 400
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(filename)[1]) as tmp_file:
            file.save(tmp_file.name)
            tmp_path = tmp_file.name
        
        try:
            if filename.endswith('.pdf'):
                text = extract_text_from_pdf(tmp_path)
            else:
                text = extract_text_from_docx(tmp_path)
        finally:
            try:
                os.unlink(tmp_path)
            except:
                pass
        
        if not text or len(text.strip()) < 50:
            return jsonify({'error': 'Could not extract text from the resume. Please ensure it contains readable text.'}), 400
        
        # Extract basic information
        name = extract_name(text)
        email = extract_email(text)
        contact = extract_phone(text)
        skills = extract_skills(text)
        experience = extract_experience(text)
        
        # Analyze job roles
        role_analysis = analyze_job_role(skills)
        
        # Get top role recommendation
        top_role = role_analysis[0] if role_analysis else None
        
        result = {
            'name': name,
            'email': email,
            'contact': contact,
            'skills': skills if skills else ["No skills found"],
            'experience': experience,
            'role_analysis': role_analysis,
            'top_role': top_role
        }
        
        return jsonify(result)
    
    except Exception as e:
        print(f"Error processing resume: {e}")
        return jsonify({'error': f'Error processing resume: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)