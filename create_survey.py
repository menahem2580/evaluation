import pandas as pd
import json
import os

# --- הגדרות ---
EXCEL_FILE = 'data.xlsx'
HTML_OUTPUT = 'index.html'
IMAGES_FOLDER = 'images'

QUESTIONS = [
    "רמת מוטיבציה", "עמידה בנהלים", "קבלת סמכות", "הבנת חומר הלימוד",
    "שואל שאלות הבנה", "יכולת למידה עצמית", "סקרנות",
    "יכולת לקבל ביקורת", "מגדיל ראש", "דירוג ביחס לכיתה"
]


def generate_html():
    print("--- מתחיל בתהליך יצירת סקר מהיר (Lazy Loading) ---")

    if not os.path.exists(EXCEL_FILE):
        print(f"❌ שגיאה: הקובץ '{EXCEL_FILE}' לא נמצא.")
        return

    try:
        df = pd.read_excel(EXCEL_FILE)
        df = df.fillna('')  # מניעת שגיאות בתאים ריקים
    except Exception as e:
        print(f"❌ שגיאה בקריאת האקסל: {e}")
        return

    students_data = []
    print(f"...קורא {len(df)} שורות מהאקסל...")

    for index, row in df.iterrows():
        try:
            # זיהוי עמודות חכם
            if 'תז' in df.columns:
                s_id = str(row['תז']).strip()
            elif 'id' in df.columns:
                s_id = str(row['id']).strip()
            else:
                # ניסיון אחרון - עמודה ראשונה אם אין שמות מוכרים
                s_id = str(row.iloc[0]).strip()

            s_name = str(row['Item Name']).strip()
            s_branch = str(row.get('סניף', 'כללי')).strip()
            s_class = str(row.get('כיתת מגמה', 'כללי')).strip()

            # טיפול בתמונה
            raw_image_name = str(row['תמונת התלמיד'])
            clean_filename = raw_image_name.replace('\\', '/').split('/')[-1]
            img_src = f"{IMAGES_FOLDER}/{clean_filename}"

            if not s_id or s_id == 'nan' or s_name == 'nan':
                continue

            students_data.append({
                "id": s_id,
                "name": s_name,
                "branch": s_branch,
                "group": s_class,
                "imagePath": img_src
            })

        except Exception as e:
            # דילוג שקט על שורות בעייתיות כדי לא לעצור את התהליך
            continue

    json_data = json.dumps(students_data, ensure_ascii=False)
    json_questions = json.dumps(QUESTIONS, ensure_ascii=False)

    print(f"✅ עובד על {len(students_data)} תלמידים. מייצר HTML מותאם לביצועים...")

    html_content = f"""
<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>מערכת סקרי מורים</title>
    <link href="https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;700&display=swap" rel="stylesheet">
    <style>
        :root {{ 
            --primary: #4a90e2; 
            --primary-dark: #357abd;
            --bg: #f5f7fa; 
            --sidebar-bg: #ffffff;
            --text-main: #333;
            --success: #2ecc71; 
            --border: #e1e4e8;
        }}

        * {{ box-sizing: border-box; }}
        body {{ 
            font-family: 'Heebo', sans-serif; 
            background: var(--bg); 
            margin: 0; 
            height: 100vh; 
            display: flex; 
            overflow: hidden; 
            color: var(--text-main);
        }}

        /* --- Sidebar --- */
        .sidebar {{ 
            width: 350px; 
            background: var(--sidebar-bg); 
            border-left: 1px solid var(--border); 
            display: flex; 
            flex-direction: column; 
            box-shadow: -2px 0 10px rgba(0,0,0,0.05);
            z-index: 10;
        }}

        .sidebar-header {{ 
            padding: 20px; 
            background: var(--primary); 
            color: white; 
            text-align: center; 
        }}
        .sidebar-header h2 {{ margin: 0; font-size: 1.4rem; font-weight: 700; }}

        .filters-box {{ 
            padding: 15px; 
            background: #f8f9fa; 
            border-bottom: 1px solid var(--border);
        }}

        .filter-label {{ font-size: 0.9rem; font-weight: bold; margin-bottom: 5px; display: block; color: #666; }}
        .filter-select {{ 
            width: 100%; 
            padding: 8px; 
            margin-bottom: 10px; 
            border: 1px solid #ccc; 
            border-radius: 6px; 
            font-family: 'Heebo', sans-serif;
            background: white;
        }}

        .actions-box {{ padding: 10px; text-align: center; border-bottom: 1px solid var(--border); }}
        .btn-export {{ 
            background: var(--success); 
            color: white; 
            border: none; 
            padding: 10px 20px; 
            width: 100%; 
            border-radius: 6px; 
            cursor: pointer; 
            font-weight: bold; 
            font-size: 1rem;
            transition: background 0.2s;
        }}
        .btn-export:hover {{ background: #27ae60; }}

        .student-list {{ overflow-y: auto; flex: 1; padding: 10px; }}

        /* שיפור ביצועי גלילה */
        .student-list {{ contain: content; }} 

        .student-card {{ 
            display: flex; 
            align-items: center; 
            padding: 12px; 
            margin-bottom: 8px; 
            background: white; 
            border: 1px solid var(--border); 
            border-radius: 8px; 
            cursor: pointer; 
            transition: all 0.2s ease;
            /* שיפור ביצועים לכרטיס */
            content-visibility: auto; 
            contain-intrinsic-size: 70px;
        }}
        .student-card:hover {{ transform: translateY(-2px); box-shadow: 0 2px 5px rgba(0,0,0,0.1); }}
        .student-card.active {{ border: 2px solid var(--primary); background: #eaf4ff; }}
        .student-card.completed {{ border-right: 4px solid var(--success); background: #f0fff4; }}
        .student-card.completed .student-info h4::after {{ content: ' ✓'; color: var(--success); font-weight: bold; }}

        .student-img {{ 
            width: 45px; 
            height: 45px; 
            border-radius: 50%; 
            object-fit: cover; 
            margin-left: 12px; 
            border: 2px solid #eee; 
            background: #ddd;
        }}
        .student-info h4 {{ margin: 0; font-size: 1rem; color: #2c3e50; }}
        .student-info p {{ margin: 0; font-size: 0.8rem; color: #7f8c8d; }}
        .student-meta {{ font-size: 0.75rem; color: #999; margin-top: 2px; }}

        /* --- Main Content --- */
        .main-content {{ 
            flex: 1; 
            padding: 30px; 
            overflow-y: auto; 
            display: flex; 
            justify-content: center; 
            align-items: flex-start;
        }}

        .form-container {{ 
            background: white; 
            padding: 40px; 
            border-radius: 12px; 
            box-shadow: 0 5px 20px rgba(0,0,0,0.08); 
            width: 100%; 
            max-width: 800px; 
            display: none; 
            animation: fadeIn 0.3s ease;
        }}

        @keyframes fadeIn {{ from {{ opacity: 0; transform: translateY(10px); }} to {{ opacity: 1; transform: translateY(0); }} }}

        .profile-header {{ 
            display: flex; 
            align-items: center; 
            margin-bottom: 30px; 
            border-bottom: 2px solid #f0f0f0; 
            padding-bottom: 20px; 
        }}
        .profile-big-img {{ 
            width: 90px; 
            height: 90px; 
            border-radius: 50%; 
            object-fit: cover; 
            margin-left: 20px; 
            border: 4px solid var(--primary); 
            box-shadow: 0 4px 10px rgba(0,0,0,0.1);
        }}
        .profile-details h1 {{ margin: 0 0 5px 0; font-size: 1.8rem; color: #2c3e50; }}
        .profile-details p {{ margin: 0; color: #7f8c8d; font-size: 1rem; }}
        .badge {{ 
            background: #eee; 
            padding: 4px 8px; 
            border-radius: 4px; 
            font-size: 0.8rem; 
            margin-right: 5px; 
            display: inline-block; 
            margin-top: 5px;
            color: #555;
        }}

        .question-row {{ 
            margin-bottom: 25px; 
            padding: 15px; 
            background: #fbfbfb; 
            border-radius: 8px; 
            border: 1px solid #eee;
        }}
        .question-label {{ font-weight: 700; font-size: 1.1rem; display: block; margin-bottom: 12px; color: #34495e; }}
        .options-wrapper {{ display: flex; gap: 8px; flex-wrap: wrap; }}

        .radio-option input[type="radio"] {{ display: none; }}
        .radio-option label {{ 
            display: flex; align-items: center; justify-content: center;
            width: 40px; height: 40px; 
            background: white; border: 1px solid #ccc; border-radius: 8px; 
            cursor: pointer; font-weight: bold; transition: all 0.2s; color: #555;
        }}
        .radio-option input:checked + label {{ 
            background: var(--primary); color: white; border-color: var(--primary); 
            transform: scale(1.1); box-shadow: 0 2px 5px rgba(74, 144, 226, 0.4);
        }}
        .radio-option.na-opt label {{ width: auto; padding: 0 15px; font-size: 0.9rem; }}
        .radio-option.na-opt input:checked + label {{ background: #95a5a6; border-color: #95a5a6; }}

        .btn-save {{ 
            background: var(--primary); color: white; padding: 15px; border: none; border-radius: 8px; width: 100%; 
            font-size: 1.2rem; font-weight: bold; cursor: pointer; margin-top: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }}
        .btn-save:hover {{ background: var(--primary-dark); }}

        .welcome-msg {{ text-align: center; margin-top: 100px; color: #bdc3c7; }}
        .welcome-msg h2 {{ font-size: 2rem; margin-bottom: 10px; }}
    </style>
</head>
<body>

<div class="sidebar">
    <div class="sidebar-header"><h2>סקר הערכת תלמידים</h2></div>
    <div class="filters-box">
        <label class="filter-label">סינון לפי סניף:</label>
        <select id="branchFilter" class="filter-select" onchange="applyFilters()"><option value="all">כל הסניפים</option></select>
        <label class="filter-label">סינון לפי כיתה:</label>
        <select id="groupFilter" class="filter-select" onchange="applyFilters()"><option value="all">כל הכיתות</option></select>
    </div>
    <div class="actions-box"><button class="btn-export" onclick="exportToCSV()">📥 הורד נתונים לאקסל</button></div>
    <div class="student-list" id="studentList"></div>
</div>

<div class="main-content">
    <div id="welcomeMsg" class="welcome-msg"><h2>👋 שלום!</h2><p>נא לבחור סניף וכיתה בצד ימין, ולאחר מכן לבחור תלמיד.</p></div>
    <form id="surveyForm" class="form-container">
        <div class="profile-header">
            <img id="currentStudentImg" class="profile-big-img" src="" onerror="this.src='https://via.placeholder.com/150'">
            <div class="profile-details">
                <h1 id="currentStudentName">שם</h1>
                <p>ת.ז: <span id="currentStudentId"></span></p>
                <div><span class="badge" id="currentBranch"></span><span class="badge" id="currentGroup"></span></div>
            </div>
        </div>
        <div id="questionsContainer"></div>
        <button type="submit" class="btn-save">שמור ועבור לתלמיד הבא</button>
    </form>
</div>

<script>
    const students = {json_data};
    const questions = {json_questions};
    let currentIndex = -1;
    let filteredStudents = [];
    let results = JSON.parse(localStorage.getItem('teacherSurveyResults_v2')) || {{}};

    function init() {{
        populateFilters();
        applyFilters();
        generateQuestionsHTML();
    }}

    function populateFilters() {{
        const branches = [...new Set(students.map(s => s.branch))].filter(b => b && b !== 'nan' && b !== '').sort();
        const groups = [...new Set(students.map(s => s.group))].filter(g => g && g !== 'nan' && g !== '').sort();

        const branchSelect = document.getElementById('branchFilter');
        branches.forEach(b => {{
            const opt = document.createElement('option'); opt.value = b; opt.textContent = b; branchSelect.appendChild(opt);
        }});

        const groupSelect = document.getElementById('groupFilter');
        groups.forEach(g => {{
            const opt = document.createElement('option'); opt.value = g; opt.textContent = g; groupSelect.appendChild(opt);
        }});
    }}

    function applyFilters() {{
        const selectedBranch = document.getElementById('branchFilter').value;
        const selectedGroup = document.getElementById('groupFilter').value;

        filteredStudents = students.filter(s => {{
            const matchBranch = (selectedBranch === 'all') || (s.branch === selectedBranch);
            const matchGroup = (selectedGroup === 'all') || (s.group === selectedGroup);
            return matchBranch && matchGroup;
        }});

        renderList();
        document.getElementById('welcomeMsg').style.display = 'block';
        document.getElementById('surveyForm').style.display = 'none';
    }}

    function renderList() {{
        const listContainer = document.getElementById('studentList');
        listContainer.innerHTML = '';

        if (filteredStudents.length === 0) {{
            listContainer.innerHTML = '<div style="text-align:center; padding:20px; color:#999">לא נמצאו תלמידים</div>';
            return;
        }}

        // כאן השינוי הגדול: שימוש ב-DocumentFragment לביצועים
        const fragment = document.createDocumentFragment();

        filteredStudents.forEach((s, index) => {{
            const isDone = results[s.id] ? 'completed' : '';
            const div = document.createElement('div');
            div.className = `student-card ${{isDone}}`;
            div.onclick = () => loadStudent(index);

            // הוספת loading="lazy" ו-decoding="async" למהירות
            div.innerHTML = `
                <img src="${{s.imagePath}}" class="student-img" loading="lazy" decoding="async" onerror="this.src='https://via.placeholder.com/50'">
                <div class="student-info">
                    <h4>${{s.name}}</h4>
                    <p>${{s.id}}</p>
                    <div class="student-meta">${{s.branch}} | ${{s.group}}</div>
                </div>
            `;
            fragment.appendChild(div);
        }});

        listContainer.appendChild(fragment);
    }}

    function generateQuestionsHTML() {{
        const container = document.getElementById('questionsContainer');
        let html = '';
        questions.forEach((q, idx) => {{
            html += `<div class="question-row"><span class="question-label">${{q}}</span><div class="options-wrapper">`;
            for(let i=1; i<=6; i++) {{
                html += `<div class="radio-option"><input type="radio" name="${{q}}" id="q${{idx}}_${{i}}" value="${{i}}" required><label for="q${{idx}}_${{i}}">${{i}}</label></div>`;
            }}
            html += `<div class="radio-option na-opt"><input type="radio" name="${{q}}" id="q${{idx}}_na" value="לא נמדד"><label for="q${{idx}}_na">לא נמדד</label></div></div></div>`;
        }});
        container.innerHTML = html;
    }}

    function loadStudent(filteredIndex) {{
        currentIndex = filteredIndex;
        const s = filteredStudents[filteredIndex];

        document.getElementById('welcomeMsg').style.display = 'none';
        document.getElementById('surveyForm').style.display = 'block';
        document.getElementById('currentStudentName').textContent = s.name;
        document.getElementById('currentStudentId').textContent = s.id;
        document.getElementById('currentBranch').textContent = s.branch;
        document.getElementById('currentGroup').textContent = s.group;

        const imgEl = document.getElementById('currentStudentImg');
        imgEl.src = s.imagePath;

        document.querySelectorAll('.student-card').forEach(c => c.classList.remove('active'));
        const activeCard = document.querySelectorAll('.student-card')[filteredIndex];
        if(activeCard) activeCard.classList.add('active');

        document.getElementById('surveyForm').reset();
        if(results[s.id]) {{
            Object.entries(results[s.id]).forEach(([k, v]) => {{
                if(k !== 'studentId' && k !== 'studentName' && k !== 'branch' && k !== 'group') {{
                    const inputs = document.getElementsByName(k);
                    inputs.forEach(input => {{ if(input.value === v) input.checked = true; }});
                }}
            }});
        }}
    }}

    document.getElementById('surveyForm').onsubmit = (e) => {{
        e.preventDefault();
        const s = filteredStudents[currentIndex];
        const formData = new FormData(e.target);
        const data = {{ studentName: s.name, studentId: s.id, branch: s.branch, group: s.group }};
        formData.forEach((v, k) => data[k] = v);

        results[s.id] = data;
        localStorage.setItem('teacherSurveyResults_v2', JSON.stringify(results));

        renderList();

        if(currentIndex < filteredStudents.length - 1) {{ loadStudent(currentIndex + 1); }} 
        else {{ alert('סיימת את כל התלמידים ברשימה זו!'); }}
    }};

    function exportToCSV() {{
        const ids = Object.keys(results);
        if(!ids.length) return alert('אין נתונים לייצוא');
        let csv = "\uFEFFשם,ת.ז,סניף,כיתה," + questions.join(",") + "\\n";
        ids.forEach(id => {{
            const r = results[id];
            let row = [`"${{r.studentName}}"` , `"${{r.studentId}}"` , `"${{r.branch}}"` , `"${{r.group}}"`];
            questions.forEach(q => row.push(`"${{r[q] || ''}}"`));
            csv += row.join(",") + "\\n";
        }});
        const link = document.createElement("a");
        link.href = URL.createObjectURL(new Blob([csv], {{type: 'text/csv;charset=utf-8;'}}));
        link.download = "תוצאות_סקר.csv";
        link.click();
    }}

    init();
</script>
</body>
</html>
    """

    with open(HTML_OUTPUT, "w", encoding='utf-8') as f:
        f.write(html_content)

    print(f"✅ הקובץ {HTML_OUTPUT} נוצר בהצלחה! כעת הוא אמור להיטען מהר יותר.")


if __name__ == "__main__":
    generate_html()