import pandas as pd
import json
import os


def generate_survey_final_mapping():
    filename = 'students.xlsx'

    # רשימת המדדים
    criteria_list = [
        "רמת מוטיבציה",
        "עמידה בנהלים",
        "קבלת סמכות",
        "הבנת חומר הלימוד",
        "שואל שאלות הבנה",
        "יכולת למידה עצמית",
        "סקרנות",
        "יכולת לקבל ביקורת",
        "מגדיל ראש",
        "דירוג ביחס לכיתה"
    ]

    if not os.path.exists(filename):
        print(f"❌ שגיאה: הקובץ {filename} לא נמצא.")
        return

    try:
        df = pd.read_excel(filename, dtype=str)
        df = df.fillna('')
        df.columns = [c.strip() for c in df.columns]
        print(f"✅ הקובץ נטען. סה'כ שורות: {len(df)}")
    except Exception as e:
        print(f"❌ שגיאה בקריאת האקסל: {e}")
        return

    # עמודות
    COL_BRANCH = 'סניף'
    COL_CLASS = 'כיתה'
    COL_NAME = 'Name'
    COL_ID = 'תז'
    COL_IMAGE = 'תמונת התלמיד'

    data_tree = {}
    total_loaded = 0

    for index, row in df.iterrows():
        # קריאת נתונים
        branch = row.get(COL_BRANCH, '').strip()
        clss = row.get(COL_CLASS, '').strip()
        name = row.get(COL_NAME, '').strip()
        st_id = row.get(COL_ID, '').strip().split('.')[0]

        if not name or name.lower() == 'nan': continue

        # --- שלב 1: תיקון זיהוי "ים" (אם כתוב בכיתה "ים" זה סניף ים) ---
        if 'ים' in clss:
            branch = 'ים'

        # --- שלב 2: החלפת השמות לפי ההוראה החדשה ---
        # "במקום פתח תקווה -> פת"
        # "במקום ים -> ירושלים"
        # "במקום ירושלים -> דרום"

        if branch == 'פתח תקווה':
            branch = 'פת'
        elif branch == 'ים':
            branch = 'ירושלים'
        elif branch == 'ירושלים':
            branch = 'דרום'
        elif branch == 'כללי' or not branch or branch.lower() == 'nan':
            branch = 'דרום'

        # הערה: אם כבר היה כתוב באקסל "דרום", זה נשאר "דרום".
        # ----------------------------------------------------

        if not clss or clss.lower() == 'nan': clss = 'ללא כיתה'

        if branch not in data_tree: data_tree[branch] = {}
        if clss not in data_tree[branch]: data_tree[branch][clss] = []

        # תמונה
        raw_img = row.get(COL_IMAGE, '').strip()
        default_img = "https://cdn-icons-png.flaticon.com/512/1077/1077114.png"

        if raw_img and raw_img.lower() != 'nan':
            filename_only = os.path.basename(raw_img.replace('\\', '/'))
            img_src = f"images/{filename_only}"
        else:
            img_src = default_img

        student_obj = {
            'name': name,
            'id': st_id,
            'image': img_src
        }
        data_tree[branch][clss].append(student_obj)
        total_loaded += 1

    print(f"✅ סה'כ תלמידים: {total_loaded}")

    # הדפסה לבדיקה
    print("\n--- הסניפים שיוצגו באתר ---")
    for b in data_tree:
        print(f"📍 {b}")

    json_data = json.dumps(data_tree, ensure_ascii=False)
    json_criteria = json.dumps(criteria_list, ensure_ascii=False)

    html = f"""
    <!DOCTYPE html>
    <html lang="he" dir="rtl">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
        <title>מערכת הערכה - קודקוד</title>
        <script src="https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js"></script>
        <link href="https://fonts.googleapis.com/css2?family=Assistant:wght@300;400;600;700&display=swap" rel="stylesheet">

        <style>
            :root {{
                --sidebar-bg: #2d3436;
                --bg-main: #f5f6fa;
                --primary: #0984e3;
                --success: #00b894;
                --danger: #d63031;
                --text: #2d3436;
            }}
            * {{ box-sizing: border-box; }}
            body {{ font-family: 'Assistant', sans-serif; background: var(--bg-main); margin: 0; display: flex; min-height: 100vh; color: var(--text); }}

            .sidebar {{
                width: 300px; background: var(--sidebar-bg); color: white; padding: 25px;
                position: fixed; right: 0; top: 0; height: 100vh; z-index: 100;
                display: flex; flex-direction: column;
                box-shadow: -5px 0 15px rgba(0,0,0,0.2);
            }}
            .brand h1 {{ margin: 0; font-size: 2.2rem; color: var(--primary); }}
            .brand span {{ opacity: 0.6; font-size: 0.8rem; letter-spacing: 1px; }}

            .filters {{ margin-top: 30px; flex-grow: 1; }}
            .filters label {{ display: block; margin-bottom: 5px; color: #b2bec3; font-weight: bold; font-size: 0.9rem; }}
            select {{ 
                width: 100%; padding: 12px; margin-bottom: 20px; border-radius: 8px; 
                border: 1px solid #636e72; background: #353b48; color: white; font-size: 1rem;
            }}

            .counter-box {{
                background: rgba(255,255,255,0.05); padding: 15px; border-radius: 8px;
                text-align: center; border: 1px solid rgba(255,255,255,0.1); margin-bottom: 20px;
            }}
            .count-num {{ font-size: 1.8rem; font-weight: bold; display: block; }}
            .count-lbl {{ font-size: 0.8rem; opacity: 0.7; }}

            .action-area {{ margin-top: auto; border-top: 1px solid rgba(255,255,255,0.1); padding-top: 20px; }}
            .excel-btn {{
                width: 100%; background: var(--success); color: white; border: none; padding: 15px;
                border-radius: 8px; font-weight: bold; cursor: pointer; display: flex; align-items: center; justify-content: center; gap: 10px;
                font-size: 1rem; transition: 0.2s;
            }}
            .excel-btn:hover {{ background: #019577; transform: translateY(-2px); }}

            .send-msg {{
                margin-top: 15px; text-align: center; font-size: 0.9rem; color: #dfe6e9;
                background: rgba(9, 132, 227, 0.2); padding: 10px; border-radius: 6px; border: 1px solid var(--primary);
            }}

            .main-content {{ margin-right: 300px; width: calc(100% - 300px); padding: 40px; }}

            .info-box {{
                background: white; border-right: 6px solid #e17055; padding: 25px; border-radius: 12px;
                margin-bottom: 40px; box-shadow: 0 4px 10px rgba(0,0,0,0.03);
            }}
            .info-box h2 {{ margin-top: 0; color: #d35400; }}
            .info-box p {{ font-size: 1.1rem; line-height: 1.6; margin-bottom: 0; color: #636e72; }}

            .student-card {{
                background: white; border-radius: 12px; margin-bottom: 30px; position: relative;
                box-shadow: 0 4px 8px rgba(0,0,0,0.03); border: 1px solid #dfe6e9; overflow: hidden;
                transition: all 0.3s;
            }}
            .student-card.completed {{ border: 2px solid var(--success); background: #f0fff4; }}
            .student-card.incomplete-alert {{ border: 2px solid var(--danger); animation: shake 0.5s; }}

            .card-header {{
                padding: 20px; background: #f8f9fa; border-bottom: 1px solid #eee;
                display: flex; align-items: center; gap: 20px;
            }}
            .s-img {{ width: 80px; height: 80px; border-radius: 50%; object-fit: cover; border: 3px solid white; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }}

            .status-badge {{
                margin-right: auto; background: var(--success); color: white; padding: 5px 15px;
                border-radius: 20px; font-weight: bold; font-size: 0.85rem; display: none;
            }}
            .student-card.completed .status-badge {{ display: block; }}

            .prog-text {{ margin-right: auto; font-size: 0.9rem; color: #e17055; font-weight: bold; }}
            .student-card.completed .prog-text {{ display: none; }}

            .criteria-grid {{ padding: 25px; display: grid; grid-template-columns: repeat(auto-fill, minmax(320px, 1fr)); gap: 25px; }}
            .c-item {{ border-bottom: 1px dashed #b2bec3; padding-bottom: 10px; }}
            .c-label {{ font-weight: 700; display: block; margin-bottom: 10px; color: #2d3436; }}
            .req {{ color: red; margin-left: 3px; }}

            .opts {{ display: flex; justify-content: space-between; gap: 5px; }}
            .rad-lbl input {{ display: none; }}

            .rad-box {{
                width: 34px; height: 34px; border-radius: 8px;
                display: flex; justify-content: center; align-items: center;
                cursor: pointer; background: #dfe6e9; color: #636e72; font-weight: bold;
                transition: 0.2s; border: 1px solid #b2bec3;
            }}

            .rad-lbl:nth-child(1) input:checked + .rad-box {{ background: #ff7675; color: white; border-color: #d63031; }}
            .rad-lbl:nth-child(2) input:checked + .rad-box {{ background: #fab1a0; color: white; border-color: #e17055; }}
            .rad-lbl:nth-child(3) input:checked + .rad-box {{ background: #fdcb6e; color: white; border-color: #e67e22; }}
            .rad-lbl:nth-child(4) input:checked + .rad-box {{ background: #ffeaa7; color: #2d3436; border-color: #fdcb6e; }}
            .rad-lbl:nth-child(5) input:checked + .rad-box {{ background: #55efc4; color: #2d3436; border-color: #00b894; }}
            .rad-lbl:nth-child(6) input:checked + .rad-box {{ background: #00b894; color: white; border-color: #00b894; }}
            .rad-lbl:last-child input:checked + .rad-box {{ background: #636e72; color: white; border-color: #2d3436; }}

            input:checked + .rad-box {{ transform: scale(1.15); box-shadow: 0 4px 10px rgba(0,0,0,0.1); }}
            .na {{ width: auto; padding: 0 10px; font-size: 0.8rem; border-radius: 20px; }}

            .notes {{ padding: 0 25px 25px; }}
            textarea {{ width: 100%; border: 1px solid #b2bec3; border-radius: 8px; padding: 15px; min-height: 80px; font-family: inherit; }}

            .hidden {{ display: none; }}
            @keyframes shake {{ 0% {{ transform: translateX(0); }} 25% {{ transform: translateX(-5px); }} 50% {{ transform: translateX(5px); }} 75% {{ transform: translateX(-5px); }} 100% {{ transform: translateX(0); }} }}

            @media (max-width: 800px) {{
                body {{ flex-direction: column; }}
                .sidebar {{ position: relative; width: 100%; height: auto; }}
                .main-content {{ margin-right: 0; width: 100%; padding: 15px; }}
            }}
        </style>
    </head>
    <body>

        <aside class="sidebar">
            <div class="brand">
                <h1>KODKOD</h1>
                <span>מערכת הערכה</span>
            </div>

            <div class="filters">
                <label>סניף</label>
                <select id="branchSelect" onchange="loadClass()">
                    <option value="">בחר סניף...</option>
                </select>

                <label>כיתה</label>
                <select id="classSelect" onchange="renderStudents()" disabled>
                    <option value="">בחר קודם סניף...</option>
                </select>

                <div class="counter-box hidden" id="counterBox">
                    <span class="count-num" id="countVal">0</span>
                    <span class="count-lbl">תלמידים בכיתה</span>
                </div>
            </div>

            <div class="action-area">
                <button onclick="validateAndDownload()" class="excel-btn hidden" id="downloadBtn">
                    <span>שמור והורד אקסל</span> 📥
                </button>
                <div class="send-msg hidden" id="sendMsg">
                    📋 נא לשלוח את הטופס ל...
                </div>
            </div>
        </aside>

        <main class="main-content">
            <div class="info-box">
                <h2>⚠️ מורים יקרים, שימו לב!</h2>
                <p>
                    למילוי טופס זה חשיבות מכרעת בעתידם של התלמידים. הנתונים שתזינו מהווים גורם משמעותי בשיבוצם ליחידות הטכנולוגיות השונות בצה"ל. 
                    הערכה כנה ומדויקת תסייע לנו להתאים לכל תלמיד את התפקיד ההולם ביותר את כישוריו, ותפתח בפניו דלתות להזדמנויות משמעותיות בשירות הצבאי ובעתיד.
                    <br><strong>תודה על שיתוף הפעולה!</strong>
                </p>
            </div>

            <div id="students-container">
                <div style="text-align:center; margin-top:50px; color:#b2bec3;">
                    <h2>אנא בחר סניף וכיתה מהתפריט</h2>
                </div>
            </div>
        </main>

        <script>
            const data = {json_data};
            const criteria = {json_criteria};

            const branchSel = document.getElementById('branchSelect');
            const classSel = document.getElementById('classSelect');
            const container = document.getElementById('students-container');
            const downloadBtn = document.getElementById('downloadBtn');
            const sendMsg = document.getElementById('sendMsg');
            const counterBox = document.getElementById('counterBox');
            const countVal = document.getElementById('countVal');

            let currentList = [];

            window.onload = () => {{
                // סדר התצוגה המדויק: פת, ירושלים, דרום
                const order = ["פת", "ירושלים", "דרום"];

                const branches = Object.keys(data);
                branches.sort((a, b) => {{
                    let indexA = order.indexOf(a);
                    let indexB = order.indexOf(b);
                    if (indexA === -1) indexA = 99;
                    if (indexB === -1) indexB = 99;
                    return indexA - indexB;
                }});

                branches.forEach(b => branchSel.add(new Option(b, b)));
            }};

            function loadClass() {{
                const branch = branchSel.value;
                classSel.innerHTML = '<option value="">בחר כיתה...</option>';
                classSel.disabled = true;
                downloadBtn.classList.add('hidden');
                sendMsg.classList.add('hidden');
                counterBox.classList.add('hidden');

                if (branch && data[branch]) {{
                    Object.keys(data[branch]).sort().forEach(c => classSel.add(new Option(c, c)));
                    classSel.disabled = false;
                }}
            }}

            function renderStudents() {{
                const branch = branchSel.value;
                const clss = classSel.value;
                container.innerHTML = '';

                if (!branch || !clss) return;

                downloadBtn.classList.remove('hidden');
                sendMsg.classList.remove('hidden');
                counterBox.classList.remove('hidden');

                const students = data[branch][clss];
                students.sort((a, b) => a.name.localeCompare(b.name));
                currentList = students;

                countVal.innerText = students.length;

                students.forEach(st => {{
                    const noteKey = `note_${{st.id}}`;

                    let criteriaHtml = '<div class="criteria-grid">';
                    criteria.forEach((crit, idx) => {{
                        const radioName = `rate_${{st.id}}_${{idx}}`;
                        let options = '';
                        for (let i = 1; i <= 6; i++) {{
                            options += `
                                <label class="rad-lbl">
                                    <input type="radio" name="${{radioName}}" value="${{i}}" onchange="checkCompletion('${{st.id}}')">
                                    <div class="rad-box">${{i}}</div>
                                </label>`;
                        }}
                        options += `
                            <label class="rad-lbl">
                                <input type="radio" name="${{radioName}}" value="NA" onchange="checkCompletion('${{st.id}}')">
                                <div class="rad-box na">לא נמדד</div>
                            </label>`;

                        criteriaHtml += `
                            <div class="c-item">
                                <span class="c-label">${{crit}}<span class="req">*</span></span>
                                <div class="opts">${{options}}</div>
                            </div>`;
                    }});
                    criteriaHtml += '</div>';

                    container.insertAdjacentHTML('beforeend', `
                        <div class="student-card" id="card_${{st.id}}">
                            <div class="card-header">
                                <img src="${{st.image}}" class="s-img">
                                <div>
                                    <h3 style="margin:0">${{st.name}}</h3>
                                    <span style="color:#636e72; font-size:0.9rem">ת.ז: ${{st.id}}</span>
                                </div>
                                <div class="status-badge">✓ הושלם</div>
                                <div class="prog-text" id="prog_${{st.id}}">ממתין למילוי...</div>
                            </div>
                            ${{criteriaHtml}}
                            <div class="notes">
                                <label style="font-weight:bold; display:block; margin-bottom:5px;">הערות:</label>
                                <textarea id="${{noteKey}}" placeholder="הערות נוספות (רשות)..." oninput="saveNote(this, '${{st.id}}')"></textarea>
                            </div>
                        </div>
                    `);

                    if (localStorage.getItem(noteKey)) document.getElementById(noteKey).value = localStorage.getItem(noteKey);
                    criteria.forEach((_, idx) => {{
                        const rName = `rate_${{st.id}}_${{idx}}`;
                        const val = localStorage.getItem(rName);
                        if (val) {{
                            const rb = document.querySelector(`input[name="${{rName}}"][value="${{val}}"]`);
                            if (rb) rb.checked = true;
                        }}
                    }});
                    checkCompletion(st.id);
                }});
            }}

            function saveNote(el, id) {{
                localStorage.setItem(el.id, el.value);
            }}

            function checkCompletion(stId) {{
                criteria.forEach((_, idx) => {{
                    const rName = `rate_${{stId}}_${{idx}}`;
                    const checked = document.querySelector(`input[name="${{rName}}"]:checked`);
                    if(checked) localStorage.setItem(rName, checked.value);
                }});

                let count = 0;
                for(let i=0; i<criteria.length; i++) {{
                    if(localStorage.getItem(`rate_${{stId}}_${{i}}`)) count++;
                }}

                const card = document.getElementById('card_'+stId);
                const prog = document.getElementById('prog_'+stId);

                if(count === criteria.length) {{
                    card.classList.add('completed');
                    card.classList.remove('incomplete-alert');
                }} else {{
                    card.classList.remove('completed');
                    prog.innerText = `מולאו ${{count}} מתוך ${{criteria.length}}`;
                }}
            }}

            function validateAndDownload() {{
                const incomplete = [];
                currentList.forEach(st => {{
                    let count = 0;
                    for(let i=0; i<criteria.length; i++) {{
                        if(localStorage.getItem(`rate_${{st.id}}_${{i}}`)) count++;
                    }}
                    if(count < criteria.length) {{
                        incomplete.push(st.name);
                        document.getElementById('card_'+st.id).classList.add('incomplete-alert');
                    }}
                }});

                if(incomplete.length > 0) {{
                    const msg = "שים לב!\\nהטפסים של התלמידים הבאים לא מולאו עד הסוף:\\n\\n" + 
                                incomplete.slice(0,5).join("\\n") + 
                                (incomplete.length > 5 ? "\\n...ועוד" : "") +
                                "\\n\\nהמערכת לא תאפשר הורדה עד להשלמת כל השדות.";
                    alert(msg);
                    return; 
                }}

                const branch = branchSel.value;
                const clss = classSel.value;
                const rows = [];

                currentList.forEach(st => {{
                    const row = {{ "שם": st.name, "תז": st.id, "סניף": branch, "כיתה": clss }};
                    criteria.forEach((crit, idx) => {{
                        row[crit] = localStorage.getItem(`rate_${{st.id}}_${{idx}}`) || "";
                    }});
                    row["הערות"] = localStorage.getItem(`note_${{st.id}}`) || "";
                    rows.push(row);
                }});

                const ws = XLSX.utils.json_to_sheet(rows);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, "דוח");
                XLSX.writeFile(wb, `דוח_${{branch}}_${{clss}}.xlsx`);
            }}
        </script>
    </body>
    </html>
    """

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)

    print("✅ הקובץ הסופי נוצר!")
    print("הסניפים מוצגים כ: פת, ירושלים, דרום.")


if __name__ == "__main__":
    generate_survey_final_mapping()