import streamlit as st
import pandas as pd
import time
from datetime import datetime, date, timedelta
import uuid
from itertools import groupby
import plotly.express as px
import os
import json

# --- KONFIGURASI APLIKASI ---
st.set_page_config(
    page_title="Safier Plan - Pengembangan Diri",
    page_icon="🌱",
    layout="wide"
)

# --- CSS Kustom ---
st.markdown("""
    <style>
        .main, [data-testid="stAppViewContainer"] { background-color: #FFF7E8; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        .stButton > button { background-color: #007BFF !important; color: white !important; font-weight: 600 !important; border-radius: 8px !important; border: none !important; padding: 0.5rem 1rem !important; margin-top: 0.5rem !important; transition: background-color 0.3s, transform 0.2s; }
        .stButton > button:hover { background-color: #0056b3 !important; transform: translateY(-2px); }
        [data-testid="stHorizontalBlock"] .stButton > button { padding: 0.25rem 0.75rem !important; font-size: 0.85rem !important; }
        button[kind="primary"] { background-color: #DC3545 !important; }
        button[kind="primary"]:hover { background-color: #c82333 !important; }
        [data-testid="stForm"], div[data-testid="stExpander"], .st-container[border="true"] { background-color: #FFFFFF !important; border: 1px solid #DEE2E6 !important; border-radius: 12px !important; padding: 1.5rem !important; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }
        [data-testid="stDateInput"] input, [data-testid="stTextInput"] input, [data-testid="stNumberInput"] input { border-radius: 8px; border: 1px solid #ced4da; padding: 0.75rem; transition: border-color 0.2s, box-shadow 0.2s; }
        [data-testid="stDateInput"] input:focus, [data-testid="stTextInput"] input:focus, [data-testid="stNumberInput"] input:focus { border-color: #007BFF; box-shadow: 0 0 0 0.2rem rgba(0,123,255,.25); }
        div[data-baseweb="tab-list"] { gap: 8px; }
        button[data-baseweb="tab"] { background-color: transparent; font-size: 1.1rem; font-weight: 600; border-radius: 8px 8px 0 0 !important; border-bottom: 2px solid transparent !important; padding: 0.5rem 1rem; }
        button[data-baseweb="tab"][aria-selected="true"] { background-color: #007BFF !important; color: white !important; border-bottom: 2px solid #0056b3 !important; }
        .table-header { background-color: #007BFF; color: white; padding: 0.75rem; border-radius: 8px; font-weight: 600; text-align: center; }
        .status-badge { padding: 0.25rem 0.6rem; border-radius: 12px; font-size: 0.8rem; font-weight: 600; color: white; text-align: center; display: inline-block; }
        .status-completed { background-color: #28a745; }
        .status-inprogress { background-color: #fd7e14; }
        .status-scheduled { background-color: #17a2b8; }
        .status-notstarted { background-color: #6c757d; }
        .area-card { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 1rem; border-radius: 12px; color: white; text-align: center; margin: 0.5rem 0; }
        .reflection-box { background-color: #F8F9FA; border-left: 4px solid #007BFF; padding: 1rem; border-radius: 8px; margin: 1rem 0; }
        .progress-label { font-size: 0.8rem; color: #666; margin-top: 0.25rem; }
    </style>
""", unsafe_allow_html=True)

# --- KONSTANTA ---
CSV_FILE = "tasks.csv"
DEVELOPMENT_FILE = "development_data.csv"
DATE_COLUMNS = ["Tanggal Mulai", "Tanggal Selesai Target", "Tanggal Jadwal", "Tanggal Selesai"]
PRIORITY_OPTIONS = ["Belum Diprioritaskan", "Lakukan Sekarang", "Jadwalkan", "Delegasikan", "Tinggalkan"]
POMODORO_DURATION = 25 * 60

# 5 Area Pengembangan Diri
DEVELOPMENT_AREAS = {
    "mental": {
        "name": "🧠 Mental",
        "icon": "🧠",
        "color": "#667eea",
        "description": "Pengembangan pikiran dan akal sehat",
        "activities": [
            "Belajar bahasa baru",
            "Membaca buku",
            "Berdialog/berdiskusi dengan orang cerdas",
            "Melakukan refleksi diri",
            "Konsultasi dengan psikolog",
            "Mengikuti workshop mental health",
            "Liburan/beristirahat",
            "Belajar keterampilan baru",
            "Memecahkan teka-teki/logika",
            "Menulis jurnal pemikiran"
        ]
    },
    "sosial": {
        "name": "🤝 Sosial",
        "icon": "🤝",
        "color": "#f093fb",
        "description": "Kemampuan berinteraksi dan berkomunikasi",
        "activities": [
            "Observasi lawan bicara",
            "Praktik active listening",
            "Latihan komunikasi efektif",
            "Belajar bahasa asing",
            "Latihan public speaking",
            "Aktif networking",
            "Memberi dan menerima feedback",
            "Bergabung dengan komunitas",
            "Volunteer/sukarelawan",
            "Membangun relasi baru"
        ]
    },
    "spiritual": {
        "name": "🙏 Spiritual",
        "icon": "🙏",
        "color": "#43e97b",
        "description": "Pengembangan nilai dan makna hidup",
        "activities": [
            "Membaca kitab suci",
            "Menghadiri kajian/ibadah",
            "Praktik self-care",
            "Meditasi/menghayati",
            "Memperhatikan lingkungan sekitar",
            "Memberi pada sesama",
            "Mempraktikkan kebaikan",
            "Bersyukur setiap hari",
            "Berkontemplasi di alam",
            "Menolong orang lain"
        ]
    },
    "emosional": {
        "name": "💖 Emosional",
        "icon": "💖",
        "color": "#fa709a",
        "description": "Pengelolaan emosi dan kecerdasan emosional",
        "activities": [
            "Menulis jurnal emosi",
            "Menjaga interaksi positif",
            "Melacak suasana hati (mood)",
            "Ekspresi diri melalui seni",
            "Berkonsultasi ke terapis",
            "Praktik mindfulness",
            "Mengelola stres",
            "Memaafkan diri dan orang lain",
            "Mengenali trigger emosi",
            "Praktik empati"
        ]
    },
    "fisik": {
        "name": "💪 Fisik",
        "icon": "💪",
        "color": "#4facfe",
        "description": "Kesehatan dan kebugaran tubuh",
        "activities": [
            "Menjaga pola makan sehat",
            "Olahraga rutin",
            "Menjaga pola tidur",
            "Membersihkan diri dan rumah",
            "Menggunakan sunscreen",
            "Menyikat gigi secara teratur",
            "Check-up kesehatan rutin",
            "Minum air putih cukup",
            "Istirahat yang cukup",
            "Menghindari rokok/alkohol"
        ]
    }
}

# --- FUNGSI CSV DATABASE ---
def init_csv():
    """Inisialisasi file CSV jika belum ada"""
    if not os.path.exists(CSV_FILE):
        df = pd.DataFrame(columns=[
            'id', 'Tugas', 'Deskripsi', 'Durasi (jam)', 
            'Tanggal Mulai', 'Tanggal Selesai Target', 
            'Selesai', 'Prioritas', 'Delegasi', 
            'Tanggal Jadwal', 'Tanggal Selesai', 'Area',
            'created_at', 'updated_at'
        ])
        df.to_csv(CSV_FILE, index=False)
    
    if not os.path.exists(DEVELOPMENT_FILE):
        df_dev = pd.DataFrame(columns=[
            'id', 'area', 'target', 'current_level', 'target_level',
            'reflection', 'action_plan', 'last_updated'
        ])
        df_dev.to_csv(DEVELOPMENT_FILE, index=False)

def load_tasks():
    """Load semua tugas dari CSV"""
    df = pd.read_csv(CSV_FILE)
    
    if 'Selesai' in df.columns:
        df['Selesai'] = df['Selesai'].fillna(False).astype(bool)
    
    for date_col in DATE_COLUMNS:
        if date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce').dt.date
    
    tasks = df.to_dict('records')
    
    for task in tasks:
        for key, value in task.items():
            if pd.isna(value):
                task[key] = None
    
    return tasks

def save_tasks(tasks):
    """Simpan semua tugas ke CSV"""
    df = pd.DataFrame(tasks)
    
    for date_col in DATE_COLUMNS:
        if date_col in df.columns:
            df[date_col] = df[date_col].apply(lambda x: x.isoformat() if pd.notna(x) else None)
    
    if 'Selesai' in df.columns:
        df['Selesai'] = df['Selesai'].astype(bool)
    
    df.to_csv(CSV_FILE, index=False)

def save_task_to_csv(task_dict):
    """Simpan atau update satu tugas ke CSV"""
    tasks = load_tasks()
    
    existing_index = None
    for i, existing_task in enumerate(tasks):
        if existing_task.get('id') == task_dict.get('id'):
            existing_index = i
            break
    
    task_dict['updated_at'] = datetime.now().isoformat()
    if 'created_at' not in task_dict or not task_dict['created_at']:
        task_dict['created_at'] = datetime.now().isoformat()
    
    if existing_index is not None:
        tasks[existing_index] = task_dict
    else:
        tasks.append(task_dict)
    
    save_tasks(tasks)
    st.cache_data.clear()

def delete_task_from_csv(task_id):
    """Hapus tugas dari CSV berdasarkan ID"""
    tasks = load_tasks()
    tasks = [task for task in tasks if task.get('id') != task_id]
    save_tasks(tasks)
    st.cache_data.clear()

def get_task_by_id(task_id):
    """Mendapatkan task berdasarkan ID"""
    tasks = load_tasks()
    return next((task for task in tasks if task.get('id') == task_id), None)

# --- FUNGSI PENGEMBANGAN DIRI ---
def load_development_data():
    """Load data pengembangan diri"""
    df = pd.read_csv(DEVELOPMENT_FILE)
    return df.to_dict('records')

def save_development_data(data):
    """Simpan data pengembangan diri"""
    df = pd.DataFrame(data)
    df.to_csv(DEVELOPMENT_FILE, index=False)

def update_development_area(area, target, current_level, target_level, reflection, action_plan):
    """Update atau create data pengembangan area tertentu"""
    data = load_development_data()
    
    existing = None
    for item in data:
        if item.get('area') == area:
            existing = item
            break
    
    new_data = {
        'id': str(uuid.uuid4()) if not existing else existing['id'],
        'area': area,
        'target': target,
        'current_level': current_level,
        'target_level': target_level,
        'reflection': reflection,
        'action_plan': action_plan,
        'last_updated': datetime.now().isoformat()
    }
    
    if existing:
        data.remove(existing)
    
    data.append(new_data)
    save_development_data(data)

def get_area_progress(area):
    """Mendapatkan progress untuk area tertentu"""
    data = load_development_data()
    for item in data:
        if item.get('area') == area:
            current = item.get('current_level', 0)
            target = item.get('target_level', 100)
            progress = (current / target * 100) if target > 0 else 0
            return progress, current, target
    return 0, 0, 100

def get_tasks_by_area(area):
    """Mendapatkan tugas berdasarkan area"""
    tasks = load_tasks()
    return [t for t in tasks if t.get('Area') == area]

# --- 3 PERTANYAAN DASAR PENGEMBANGAN DIRI ---
def render_three_questions():
    """Render 3 pertanyaan dasar pengembangan diri"""
    st.markdown("### 📋 3 Pertanyaan Dasar Pengembangan Diri")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        with st.container(border=True):
            st.markdown("#### 🎯 Where do I want to be?")
            st.caption("Di mana saya ingin berada?")
            if 'vision_goal' not in st.session_state:
                st.session_state.vision_goal = ""
            vision = st.text_area("Tulis visi Anda:", 
                                 value=st.session_state.vision_goal,
                                 placeholder="Contoh: Dalam 1 tahun ke depan, saya ingin menjadi pribadi yang lebih percaya diri dan mampu berkomunikasi dengan baik...",
                                 key="vision_input",
                                 height=150)
            st.session_state.vision_goal = vision
    
    with col2:
        with st.container(border=True):
            st.markdown("#### 📍 Where am I now?")
            st.caption("Di mana saya sekarang?")
            if 'current_state' not in st.session_state:
                st.session_state.current_state = ""
            current = st.text_area("Evaluasi diri Anda saat ini:",
                                  value=st.session_state.current_state,
                                  placeholder="Contoh: Saat ini saya masih merasa grogi saat presentasi dan kurang percaya diri...",
                                  key="current_input",
                                  height=150)
            st.session_state.current_state = current
    
    with col3:
        with st.container(border=True):
            st.markdown("#### 🗺️ How do I get there?")
            st.caption("Bagaimana saya menuju ke sana?")
            if 'action_strategy' not in st.session_state:
                st.session_state.action_strategy = ""
            strategy = st.text_area("Rencana aksi Anda:",
                                   value=st.session_state.action_strategy,
                                   placeholder="Contoh: Saya akan mengikuti kursus public speaking, praktik setiap hari, dan meminta feedback...",
                                   key="strategy_input",
                                   height=150)
            st.session_state.action_strategy = strategy
    
    col_save1, col_save2, col_save3 = st.columns(3)
    if col_save2.button("💾 Simpan Visi Pengembangan Diri", use_container_width=True):
        st.success("✅ Visi dan rencana pengembangan diri berhasil disimpan!")
        # Simpan ke session state sudah otomatis

# --- KOMPONEN UI ---
def render_metrics(tasks_aktif, tasks_selesai_hari_ini, tasks_mendesak):
    """Render metrics cards"""
    col1, col2, col3 = st.columns(3)
    col1.metric(label="✅ Tugas Selesai Hari Ini", value=len(tasks_selesai_hari_ini))
    col2.metric(label="📋 Total Tugas Aktif", value=len(tasks_aktif))
    col3.metric(label="⚠️ Tugas Mendesak", value=len(tasks_mendesak))

def render_development_dashboard():
    """Dashboard 5 Area Pengembangan Diri"""
    st.header("🌱 Dashboard 5 Area Pengembangan Diri")
    
    # Tampilkan 5 area dalam grid
    cols = st.columns(5)
    for idx, (area_key, area_info) in enumerate(DEVELOPMENT_AREAS.items()):
        with cols[idx]:
            progress, current, target = get_area_progress(area_key)
            st.markdown(f"""
            <div class="area-card" style="background: linear-gradient(135deg, {area_info['color']} 0%, {area_info['color']}cc 100%);">
                <h2>{area_info['icon']}</h2>
                <h4>{area_info['name']}</h4>
                <div class="progress-label">Progress: {progress:.0f}%</div>
            </div>
            """, unsafe_allow_html=True)
            
            if st.button(f"Detail {area_info['icon']}", key=f"detail_{area_key}", use_container_width=True):
                st.session_state.selected_area = area_key
                st.rerun()
    
    st.markdown("---")
    
    # Tampilkan detail area yang dipilih
    if 'selected_area' in st.session_state and st.session_state.selected_area:
        area_key = st.session_state.selected_area
        area_info = DEVELOPMENT_AREAS[area_key]
        
        st.subheader(f"{area_info['icon']} {area_info['name']}")
        st.markdown(f"*{area_info['description']}*")
        
        # Load existing data
        dev_data = load_development_data()
        area_data = next((d for d in dev_data if d.get('area') == area_key), None)
        
        # Form untuk pengembangan area
        with st.form(f"form_{area_key}"):
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**🎯 Target Pengembangan**")
                target = st.text_area(
                    "Apa target spesifik Anda?",
                    value=area_data.get('target', '') if area_data else '',
                    placeholder=f"Contoh: Meningkatkan kemampuan {DEVELOPMENT_AREAS[area_key]['name'].lower()}...",
                    key=f"target_{area_key}"
                )
                
                st.markdown("**📊 Level Saat Ini**")
                current_level = st.slider(
                    "Level saat ini (0-100):",
                    min_value=0, max_value=100,
                    value=int(area_data.get('current_level', 0)) if area_data else 0,
                    key=f"current_{area_key}"
                )
                
                st.markdown("**🎯 Level Target**")
                target_level = st.slider(
                    "Level target (0-100):",
                    min_value=0, max_value=100,
                    value=int(area_data.get('target_level', 80)) if area_data else 80,
                    key=f"target_level_{area_key}"
                )
            
            with col2:
                st.markdown("**🤔 Refleksi Diri**")
                reflection = st.text_area(
                    "Apa yang sudah dan belum baik?",
                    value=area_data.get('reflection', '') if area_data else '',
                    placeholder="Refleksikan kondisi Anda saat ini...",
                    height=150,
                    key=f"reflection_{area_key}"
                )
                
                st.markdown("**📝 Rencana Aksi**")
                action_plan = st.text_area(
                    "Apa yang akan Anda lakukan?",
                    value=area_data.get('action_plan', '') if area_data else '',
                    placeholder="Buat rencana konkret...",
                    height=150,
                    key=f"action_{area_key}"
                )
            
            st.markdown("---")
            st.markdown("**💡 Rekomendasi Aktivitas:**")
            activities = area_info['activities'][:5]
            activity_cols = st.columns(len(activities))
            for idx, activity in enumerate(activities):
                with activity_cols[idx]:
                    st.info(f"📌 {activity}")
            
            if st.form_submit_button(f"💾 Simpan Progress {area_info['name']}", use_container_width=True):
                update_development_area(area_key, target, current_level, target_level, reflection, action_plan)
                st.success(f"✅ Progress {area_info['name']} berhasil disimpan!")
                st.rerun()
        
        # Tampilkan tugas terkait area ini
        st.markdown("---")
        st.subheader("📋 Tugas Terkait Area Ini")
        
        area_tasks = get_tasks_by_area(area_key)
        if area_tasks:
            for task in area_tasks:
                status_text, _ = get_task_status(task)
                col_check, col_task = st.columns([0.1, 0.9])
                with col_check:
                    is_done = st.checkbox("", value=task.get('Selesai'), key=f"area_task_{task['id']}")
                    if is_done != task.get('Selesai'):
                        task['Selesai'] = is_done
                        task['Tanggal Selesai'] = date.today() if is_done else None
                        save_task_to_csv(task)
                        st.rerun()
                with col_task:
                    st.write(f"**{task.get('Tugas')}** - {status_text}")
        else:
            st.info(f"Belum ada tugas untuk area {area_info['name']}. Buat tugas baru dan pilih area ini!")
        
        if st.button("🔙 Kembali ke Ringkasan", use_container_width=True):
            del st.session_state.selected_area
            st.rerun()

def render_task_table(tasks):
    """Render tabel tugas dengan optimasi"""
    if not tasks:
        st.warning("Tidak ada tugas yang cocok dengan kriteria pencarian Anda.")
        return
    
    header_cols = st.columns((2, 2, 1.2, 1.2, 1.2, 1.2, 1.2, 1.5))
    col_names = ["Tugas", "Deskripsi", "Area", "Status", "Tgl Mulai", "Tgl Selesai", "Prioritas", "Aksi"]
    for col, name in zip(header_cols, col_names):
        col.markdown(f'<p class="table-header">{name}</p>', unsafe_allow_html=True)
    
    for task in tasks:
        st.markdown("---")
        row_cols = st.columns((2, 2, 1.2, 1.2, 1.2, 1.2, 1.2, 1.5))
        
        status_text, status_class = get_task_status(task)
        area_name = DEVELOPMENT_AREAS.get(task.get('Area', ''), {}).get('name', '-') if task.get('Area') else '-'
        
        row_cols[0].write(task.get('Tugas'))
        row_cols[1].write(task.get('Deskripsi', '')[:80] + ('...' if len(task.get('Deskripsi', '')) > 80 else ''))
        row_cols[2].write(area_name)
        row_cols[3].markdown(f'<div class="status-badge {status_class}">{status_text}</div>', unsafe_allow_html=True)
        row_cols[4].write(task.get('Tanggal Mulai').strftime('%d %b %Y') if task.get('Tanggal Mulai') else "-")
        row_cols[5].write(task.get('Tanggal Selesai Target').strftime('%d %b %Y') if task.get('Tanggal Selesai Target') else "-")
        row_cols[6].write(task.get('Prioritas'))
        
        with row_cols[7]:
            action_cols = st.columns(2)
            if action_cols[0].button("✏️", key=f"edit_{task['id']}", use_container_width=True):
                st.session_state.editing_task_id = task['id']
                st.rerun()
            if action_cols[1].button("🗑️", key=f"del_{task['id']}", type="primary", use_container_width=True):
                delete_task_from_csv(task['id'])
                st.success(f"Tugas '{task.get('Tugas')}' berhasil dihapus.")
                st.rerun()

def render_new_task_form():
    """Render form untuk menambah tugas baru dengan area pengembangan diri"""
    with st.expander("➕ Tambah Tugas Baru", expanded=False):
        with st.form("new_task_form", clear_on_submit=True):
            st.subheader("📝 Detail Tugas Utama")
            
            task_input = st.text_input("Nama Tugas:", placeholder="Contoh: Membaca buku pengembangan diri 30 menit")
            
            col1, col2 = st.columns(2)
            start_date = col1.date_input("📅 Tanggal Mulai", value=date.today())
            end_date = col2.date_input("🎯 Tanggal Selesai Target", value=date.today() + timedelta(days=7))
            
            st.subheader("🏷️ Kategorisasi")
            col_area1, col_area2 = st.columns(2)
            with col_area1:
                area_options = ["", *[f"{info['icon']} {info['name']}" for info in DEVELOPMENT_AREAS.values()]]
                selected_area_display = st.selectbox("Area Pengembangan Diri:", options=area_options)
                selected_area_key = next((key for key, info in DEVELOPMENT_AREAS.items() 
                                         if f"{info['icon']} {info['name']}" == selected_area_display), None)
            
            with col_area2:
                priority = st.selectbox("Prioritas:", options=PRIORITY_OPTIONS)
            
            st.subheader("🎯 Definisi Tugas SMART (Opsional)")
            smart_s = st.text_input("S (Specific):", placeholder="Apa yang ingin dicapai?")
            smart_m = st.text_input("M (Measurable):", placeholder="Bagaimana mengukurnya?")
            smart_a = st.text_input("A (Achievable):", placeholder="Apakah realistis?")
            smart_r = st.text_input("R (Relevant):", placeholder="Mengapa penting?")
            smart_t = st.text_input("T (Time-bound):", placeholder="Kapan selesainya?")
            
            submitted = st.form_submit_button("💾 Simpan Tugas", use_container_width=True)
            
            if submitted and task_input:
                smart_parts = [s for s in [smart_s, smart_m, smart_a, smart_r, smart_t] if s]
                full_description = "\n".join([f"• {part}" for part in smart_parts]) if smart_parts else "Tidak ada deskripsi."
                
                new_task = {
                    "id": str(uuid.uuid4()),
                    "Tugas": task_input,
                    "Deskripsi": full_description,
                    "Durasi (jam)": 0.0,
                    "Tanggal Mulai": start_date,
                    "Tanggal Selesai Target": end_date,
                    "Selesai": False,
                    "Prioritas": priority,
                    "Delegasi": "",
                    "Tanggal Jadwal": None,
                    "Tanggal Selesai": None,
                    "Area": selected_area_key,
                    "created_at": datetime.now().isoformat(),
                    "updated_at": datetime.now().isoformat()
                }
                
                save_task_to_csv(new_task)
                st.success(f"✅ Tugas '{task_input}' berhasil ditambahkan!")
                st.rerun()
            elif submitted:
                st.warning("⚠️ Nama tugas tidak boleh kosong.")

def render_edit_task_form():
    """Render form untuk mengedit tugas"""
    task_to_edit = get_task_by_id(st.session_state.editing_task_id)
    if not task_to_edit:
        st.session_state.editing_task_id = None
        return
    
    with st.form("edit_task_form"):
        st.subheader(f"✏️ Edit Tugas: {task_to_edit.get('Tugas', '')}")
        
        new_task_name = st.text_input("Nama Tugas", value=task_to_edit.get('Tugas', ''))
        new_desc = st.text_area("Deskripsi", value=task_to_edit.get('Deskripsi', ''), height=150)
        
        col1, col2 = st.columns(2)
        new_start_date = col1.date_input("Tanggal Mulai", value=task_to_edit.get('Tanggal Mulai', date.today()))
        new_end_date = col2.date_input("Tanggal Selesai Target", value=task_to_edit.get('Tanggal Selesai Target', date.today()))
        
        area_options = ["", *[f"{info['icon']} {info['name']}" for info in DEVELOPMENT_AREAS.values()]]
        current_area_display = ""
        if task_to_edit.get('Area'):
            area_info = DEVELOPMENT_AREAS.get(task_to_edit['Area'], {})
            current_area_display = f"{area_info.get('icon', '')} {area_info.get('name', '')}"
        
        selected_area_display = st.selectbox("Area Pengembangan Diri:", options=area_options, index=area_options.index(current_area_display) if current_area_display in area_options else 0)
        selected_area_key = next((key for key, info in DEVELOPMENT_AREAS.items() 
                                 if f"{info['icon']} {info['name']}" == selected_area_display), None)
        
        priority = st.selectbox("Prioritas:", options=PRIORITY_OPTIONS, index=PRIORITY_OPTIONS.index(task_to_edit.get('Prioritas', 'Belum Diprioritaskan')) if task_to_edit.get('Prioritas') in PRIORITY_OPTIONS else 0)
        
        col_btn1, col_btn2 = st.columns(2)
        if col_btn1.form_submit_button("💾 Simpan Perubahan", use_container_width=True):
            task_to_edit.update({
                'Tugas': new_task_name,
                'Deskripsi': new_desc,
                'Tanggal Mulai': new_start_date,
                'Tanggal Selesai Target': new_end_date,
                'Area': selected_area_key,
                'Prioritas': priority,
                'updated_at': datetime.now().isoformat()
            })
            save_task_to_csv(task_to_edit)
            st.session_state.editing_task_id = None
            st.success("✅ Tugas berhasil diperbarui!")
            st.rerun()
        
        if col_btn2.form_submit_button("❌ Batal", use_container_width=True):
            st.session_state.editing_task_id = None
            st.rerun()

def get_task_status(task):
    """Mendapatkan status dan warna task"""
    if task.get('Selesai'):
        return "Completed", "status-completed"
    if task.get('Prioritas') == 'Lakukan Sekarang':
        return "In Progress", "status-inprogress"
    if task.get('Prioritas') == 'Jadwalkan':
        return "Scheduled", "status-scheduled"
    return "Not Started", "status-notstarted"

def check_and_update_due_tasks():
    """Cek dan update tugas yang sudah waktunya dikerjakan"""
    tasks = load_tasks()
    today_date = date.today()
    updated = False
    
    for task in tasks:
        if (task.get('Prioritas') == 'Jadwalkan' and 
            task.get('Tanggal Jadwal') and 
            isinstance(task.get('Tanggal Jadwal'), date) and 
            task.get('Tanggal Jadwal') == today_date and
            not task.get('Selesai')):
            task['Prioritas'] = 'Lakukan Sekarang'
            save_task_to_csv(task)
            updated = True
            st.toast(f"✨ Tugas '{task.get('Tugas')}' kini menjadi 'Lakukan Sekarang'!")
    
    if updated:
        st.rerun()

def get_filtered_tasks(tasks, search_query, selected_priorities, selected_area):
    """Filter tasks berdasarkan query, prioritas, dan area"""
    filtered = tasks
    
    if search_query:
        filtered = [t for t in filtered if search_query.lower() in t.get('Tugas', '').lower()]
    
    if selected_priorities:
        filtered = [t for t in filtered if t.get('Prioritas') in selected_priorities]
    
    if selected_area:
        filtered = [t for t in filtered if t.get('Area') == selected_area]
    
    return filtered

def render_development_report():
    """Laporan perkembangan 5 area"""
    st.subheader("📊 Laporan Perkembangan 5 Area")
    
    dev_data = load_development_data()
    
    if not dev_data:
        st.info("Belum ada data perkembangan. Mulai isi progres di tab '5 Area Pengembangan Diri'!")
        return
    
    # Progress chart untuk semua area
    progress_data = []
    for area_key, area_info in DEVELOPMENT_AREAS.items():
        progress, current, target = get_area_progress(area_key)
        if current > 0 or target > 0:
            progress_data.append({
                'Area': area_info['name'],
                'Progress': progress,
                'Current': current,
                'Target': target
            })
    
    if progress_data:
        df_progress = pd.DataFrame(progress_data)
        fig = px.bar(
            df_progress,
            x='Area',
            y='Progress',
            title='Progress Pengembangan Diri per Area',
            text='Progress',
            color='Progress',
            color_continuous_scale='Viridis',
            range_y=[0, 100]
        )
        fig.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
        st.plotly_chart(fig, use_container_width=True)
    
    st.markdown("---")
    
    # Detail per area
    st.subheader("📋 Detail Per Area")
    for area_key, area_info in DEVELOPMENT_AREAS.items():
        area_data = next((d for d in dev_data if d.get('area') == area_key), None)
        if area_data:
            with st.expander(f"{area_info['icon']} {area_info['name']}"):
                progress, current, target = get_area_progress(area_key)
                st.metric("Progress", f"{progress:.1f}%", f"{current}/{target}")
                st.progress(progress/100)
                
                if area_data.get('target'):
                    st.markdown("**🎯 Target:**")
                    st.write(area_data['target'])
                
                if area_data.get('reflection'):
                    st.markdown("**🤔 Refleksi:**")
                    st.write(area_data['reflection'])
                
                if area_data.get('action_plan'):
                    st.markdown("**📝 Rencana Aksi:**")
                    st.write(area_data['action_plan'])
                
                # Tugas terkait area ini
                area_tasks = get_tasks_by_area(area_key)
                if area_tasks:
                    st.markdown("**📋 Tugas terkait:**")
                    for task in area_tasks:
                        status = "✅" if task.get('Selesai') else "⏳"
                        st.write(f"{status} {task.get('Tugas')}")

# --- INISIALISASI ---
init_csv()

# Inisialisasi session state
if 'tasks' not in st.session_state:
    st.session_state.tasks = load_tasks()
if 'editing_task_id' not in st.session_state:
    st.session_state.editing_task_id = None
if 'pomodoro_running' not in st.session_state:
    st.session_state.pomodoro_running = False
if 'active_pomodoro_task' not in st.session_state:
    st.session_state.active_pomodoro_task = None
if 'pomodoro_start_time' not in st.session_state:
    st.session_state.pomodoro_start_time = 0
if 'vision_goal' not in st.session_state:
    st.session_state.vision_goal = ""
if 'current_state' not in st.session_state:
    st.session_state.current_state = ""
if 'action_strategy' not in st.session_state:
    st.session_state.action_strategy = ""

# Refresh tasks dari CSV
st.session_state.tasks = load_tasks()

# Cek dan update tugas yang sudah waktunya
check_and_update_due_tasks()

# --- DATA UNTUK METRICS ---
tasks_aktif = [t for t in st.session_state.tasks if not t.get('Selesai')]
tasks_selesai_hari_ini = [t for t in st.session_state.tasks if t.get('Selesai') and t.get('Tanggal Selesai') == date.today()]
tasks_mendesak = [t for t in tasks_aktif if t.get('Prioritas') == 'Lakukan Sekarang']

# --- TAMPILAN UTAMA ---
st.title("🌱 Safier Plan - Pengembangan Diri Holistik")
st.markdown("*Integrasi Manajemen Tugas dengan 5 Area Pengembangan Diri: Mental, Sosial, Spiritual, Emosional, dan Fisik*")

# Sidebar
with st.sidebar:
    st.header("⚙️ Menu Utama")
    
    # Informasi pengguna
    st.markdown("### 🎯 Visi Pengembangan Diri")
    if st.session_state.vision_goal:
        with st.expander("Lihat Visi Saya"):
            st.write(st.session_state.vision_goal)
    
    st.markdown("---")
    
    # Manajemen data
    st.subheader("💾 Manajemen Data")
    if os.path.exists(CSV_FILE):
        file_size = os.path.getsize(CSV_FILE) / 1024
        st.info(f"📁 Database: {CSV_FILE}\n📊 Ukuran: {file_size:.2f} KB\n📝 Total tugas: {len(st.session_state.tasks)}")
    
    if st.button("📀 Backup Data", use_container_width=True):
        import shutil
        backup_name = f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{CSV_FILE}"
        shutil.copy(CSV_FILE, backup_name)
        st.success(f"✅ Backup: {backup_name}")
    
    st.markdown("---")
    st.caption("💡 *Kembangkan diri secara seimbang di 5 area*")

# Header dengan metrics
st.header("📊 Ringkasan Produktivitas", divider='rainbow')
render_metrics(tasks_aktif, tasks_selesai_hari_ini, tasks_mendesak)

# Tabs utama
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9 = st.tabs([
    "🏠 Beranda",
    "📝 Tugas", 
    "🎯 Prioritaskan", 
    "🌱 5 Area Pengembangan", 
    "📅 Jadwal", 
    "🗓️ Kalender", 
    "⏱️ Sesi Fokus", 
    "🤝 Delegasi", 
    "📈 Laporan"
])

# --- TAB 1: BERANDA ---
with tab1:
    st.header("🏠 Selamat Datang di Safier Plan")
    
    st.markdown("""
    ### 🌟 Filosofi Pengembangan Diri
    
    > *"Bakat saja tidak cukup. Keahlian adalah hasil dari proses belajar yang terus-menerus. 
    > Proses belajar yang dimaksud bukan soal membuat otak encer saja, tapi juga soal kegigihan, komunikasi, bahkan kolaborasi."*
    
    ### 📌 5 Area Pengembangan Diri yang Seimbang
    
    Seperti yang telah kita pelajari, pengembangan diri yang optimal mencakup 5 area:
    
    | Area | Fokus | Aktivitas Utama |
    |------|-------|-----------------|
    | 🧠 **Mental** | Pikiran dan akal sehat | Membaca, belajar, diskusi, refleksi |
    | 🤝 **Sosial** | Interaksi dan komunikasi | Active listening, networking, public speaking |
    | 🙏 **Spiritual** | Nilai dan makna hidup | Meditasi, self-care, berbagi dengan sesama |
    | 💖 **Emosional** | Pengelolaan emosi | Jurnal, mindfulness, manajemen stres |
    | 💪 **Fisik** | Kesehatan tubuh | Olahraga, pola makan, istirahat |
    
    ### 🎯 3 Pertanyaan Dasar
    
    Sebelum memulai, jawablah 3 pertanyaan ini:
    """)
    
    render_three_questions()
    
    st.markdown("---")
    
    # Tampilan ringkasan 5 area
    st.subheader("📊 Ringkasan Progress 5 Area")
    cols = st.columns(5)
    for idx, (area_key, area_info) in enumerate(DEVELOPMENT_AREAS.items()):
        with cols[idx]:
            progress, _, _ = get_area_progress(area_key)
            st.markdown(f"""
            <div class="area-card" style="background: linear-gradient(135deg, {area_info['color']} 0%, {area_info['color']}cc 100%); padding: 1rem; border-radius: 12px; text-align: center;">
                <h2>{area_info['icon']}</h2>
                <h5>{area_info['name']}</h5>
                <h3>{progress:.0f}%</h3>
            </div>
            """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Tips pengembangan diri
    with st.expander("💡 Tips Pengembangan Diri Harian", expanded=False):
        st.markdown("""
        **Rutinitas 15 menit setiap hari:**
        - 🧠 **Mental** (3 menit): Baca 1 halaman buku
        - 🤝 **Sosial** (3 menit): Kirim pesan positif ke teman
        - 🙏 **Spiritual** (3 menit): Refleksi/meditasi singkat
        - 💖 **Emosional** (3 menit): Tulis 1 hal yang disyukuri
        - 💪 **Fisik** (3 menit): Stretching ringan
        
        **Ingat:** Konsistensi lebih penting daripada intensitas!
        """)

# --- TAB 2: TUGAS ---
with tab2:
    st.header("📝 Input & Kelola Semua Tugas")
    
    if st.session_state.editing_task_id:
        render_edit_task_form()
        st.markdown("---")
    
    render_new_task_form()
    
    st.markdown("---")
    st.subheader("📋 Daftar Semua Tugas")
    
    # Filter
    col_filter1, col_filter2, col_filter3 = st.columns([2, 1, 1])
    with col_filter1:
        search_query = st.text_input("🔍 Cari tugas:", placeholder="Ketik nama tugas...")
    with col_filter2:
        selected_priorities = st.multiselect("🏷️ Prioritas:", options=PRIORITY_OPTIONS)
    with col_filter3:
        area_options = ["Semua Area", *[info['name'] for info in DEVELOPMENT_AREAS.values()]]
        selected_area_display = st.selectbox("🌱 Area:", options=area_options)
        selected_area = next((key for key, info in DEVELOPMENT_AREAS.items() 
                             if info['name'] == selected_area_display), None) if selected_area_display != "Semua Area" else None
    
    filtered_tasks = get_filtered_tasks(st.session_state.tasks, search_query, selected_priorities, selected_area)
    render_task_table(filtered_tasks)

# --- TAB 3: PRIORITASKAN ---
with tab3:
    st.header("🎯 Prioritaskan Tugas Anda")
    st.markdown("""
    ### 📌 Matriks Prioritas Eisenhower
    
    | Kuadran | Prioritas | Tindakan | Contoh |
    |---------|-----------|----------|--------|
    | **I** | Lakukan Sekarang | Penting & Mendesak | Deadline dekat, krisis |
    | **II** | Jadwalkan | Penting & Tidak Mendesak | Perencanaan, belajar |
    | **III** | Delegasikan | Tidak Penting & Mendesak | Tugas administratif |
    | **IV** | Tinggalkan | Tidak Penting & Tidak Mendesak | Distraksi |
    """)
    
    st.markdown("---")
    
    # Daftar tugas aktif
    st.subheader("📋 Daftar Tugas Aktif")
    if not tasks_aktif:
        st.info("✨ Tidak ada tugas aktif untuk diprioritaskan. Selamat!")
    else:
        df_aktif = pd.DataFrame(tasks_aktif)
        df_display = df_aktif[['Tugas', 'Prioritas', 'Area', 'Tanggal Mulai', 'Tanggal Selesai Target']]
        df_display.columns = ['Tugas', 'Prioritas', 'Area', 'Tanggal Mulai', 'Target Selesai']
        st.dataframe(df_display, use_container_width=True, hide_index=True)
    
    st.markdown("---")
    
    # Bank tugas
    st.info("💡 Pilih tugas dari 'Bank Tugas' dan klik tombol kuadran untuk memprioritaskan.")
    
    bank_tugas = [task for task in tasks_aktif if task.get('Prioritas') == 'Belum Diprioritaskan']
    
    if bank_tugas:
        selected_task_name = st.selectbox(
            "Pilih Tugas:", 
            options=[t['Tugas'] for t in bank_tugas],
            key="task_to_prioritize"
        )
        selected_task = next((t for t in bank_tugas if t.get('Tugas') == selected_task_name), None)
        
        if selected_task:
            st.subheader("📊 Pindahkan ke Kuadran:")
            
            col_q1, col_q2, col_q3, col_q4 = st.columns(4)
            
            with col_q1:
                if st.button("🔥 Lakukan Sekarang", use_container_width=True):
                    update_task_priority_global(selected_task['id'], "Lakukan Sekarang")
                    st.success(f"✅ Tugas dipindahkan ke 'Lakukan Sekarang'")
                    st.rerun()
            
            with col_q2:
                if st.button("📅 Jadwalkan", use_container_width=True):
                    update_task_priority_global(selected_task['id'], "Jadwalkan")
                    st.success(f"✅ Tugas dipindahkan ke 'Jadwalkan'")
                    st.rerun()
            
            with col_q3:
                if st.button("🤝 Delegasikan", use_container_width=True):
                    update_task_priority_global(selected_task['id'], "Delegasikan")
                    st.success(f"✅ Tugas dipindahkan ke 'Delegasikan'")
                    st.rerun()
            
            with col_q4:
                if st.button("🗑️ Tinggalkan", use_container_width=True):
                    update_task_priority_global(selected_task['id'], "Tinggalkan")
                    st.warning(f"⚠️ Tugas dipindahkan ke 'Tinggalkan'")
                    st.rerun()
    else:
        st.success("🎉 Semua tugas aktif sudah diprioritaskan!")

def update_task_priority_global(task_id, new_priority):
    """Update prioritas task"""
    task = get_task_by_id(task_id)
    if task:
        task['Prioritas'] = new_priority
        save_task_to_csv(task)

# --- TAB 4: 5 AREA PENGEMBANGAN ---
with tab4:
    render_development_dashboard()

# --- TAB 5: JADWAL AKTIVITAS ---
with tab5:
    st.header("📅 Jadwal Aktivitas Anda")
    
    tasks_to_schedule = sorted(
        [t for t in tasks_aktif if t.get('Prioritas') in ["Lakukan Sekarang", "Jadwalkan"]],
        key=lambda x: x.get('Tanggal Jadwal') or x.get('Tanggal Mulai') or date.max
    )
    
    if not tasks_to_schedule:
        st.info("📭 Tidak ada tugas yang perlu dijadwalkan.")
    else:
        st.success("✅ Tandai checklist untuk mencatat tugas yang sudah selesai!")
        
        grouped_tasks = {}
        for task in tasks_to_schedule:
            task_date = task.get('Tanggal Jadwal') or task.get('Tanggal Mulai')
            if task_date:
                if task_date not in grouped_tasks:
                    grouped_tasks[task_date] = []
                grouped_tasks[task_date].append(task)
        
        for task_date in sorted(grouped_tasks.keys()):
            nice_date = task_date.strftime("%A, %d %B %Y")
            is_today = task_date == date.today()
            
            with st.container(border=True):
                if is_today:
                    st.subheader(f"🌟 HARI INI - {nice_date}")
                else:
                    st.subheader(f"📌 {nice_date}")
                
                for task in grouped_tasks[task_date]:
                    col_check, col_task = st.columns([0.1, 0.9])
                    
                    with col_check:
                        is_completed = st.checkbox(
                            "",
                            value=task.get('Selesai', False),
                            key=f"complete_{task['id']}",
                            label_visibility="collapsed"
                        )
                        
                        if is_completed != task.get('Selesai'):
                            task['Selesai'] = is_completed
                            task['Tanggal Selesai'] = date.today() if is_completed else None
                            save_task_to_csv(task)
                            st.rerun()
                    
                    with col_task:
                        area_icon = DEVELOPMENT_AREAS.get(task.get('Area'), {}).get('icon', '📋')
                        priority_emoji = "🔥" if task.get('Prioritas') == 'Lakukan Sekarang' else "📅"
                        st.write(f"{area_icon} {priority_emoji} **{task.get('Tugas')}**")
                        if task.get('Deskripsi'):
                            st.caption(task.get('Deskripsi', '')[:100])

# --- TAB 6: KALENDER ---
with tab6:
    st.header("🗓️ Visualisasi Kalender Tugas")
    
    calendar_events = []
    for task in st.session_state.tasks:
        start_date = task.get("Tanggal Mulai")
        end_date = task.get("Tanggal Selesai Target")
        
        if start_date and end_date:
            _, status_class = get_task_status(task)
            color_map = {
                "status-completed": "#28a745",
                "status-inprogress": "#fd7e14",
                "status-scheduled": "#17a2b8",
                "status-notstarted": "#6c757d"
            }
            
            calendar_events.append({
                "title": f"{DEVELOPMENT_AREAS.get(task.get('Area'), {}).get('icon', '📋')} {task['Tugas']}",
                "start": start_date.isoformat(),
                "end": (end_date + timedelta(days=1)).isoformat(),
                "color": color_map.get(status_class, "#007BFF")
            })
    
    if calendar_events:
        from streamlit_calendar import calendar
        calendar_options = {
            "headerToolbar": {
                "left": "prev,next today",
                "center": "title",
                "right": "dayGridMonth,timeGridWeek,timeGridDay"
            },
            "initialView": "dayGridMonth",
            "height": "700px",
            "buttonText": {
                "today": "Hari Ini",
                "month": "Bulan",
                "week": "Minggu",
                "day": "Hari"
            }
        }
        calendar(events=calendar_events, options=calendar_options, key="main_calendar")
    else:
        st.info("📭 Tidak ada tugas dengan tanggal yang bisa ditampilkan di kalender.")

# --- TAB 7: SESI FOKUS ---
with tab7:
    st.header("⏱️ Sesi Fokus (Pomodoro)")
    st.markdown("""
    ### 🍅 Teknik Pomodoro untuk Pengembangan Diri
    - 25 menit fokus penuh pada satu tugas pengembangan diri
    - 5 menit istirahat untuk refleksi singkat
    - Setelah 4 sesi, istirahat 15-30 menit
    """)
    
    if st.session_state.pomodoro_running:
        elapsed_time = time.time() - st.session_state.pomodoro_start_time
        time_left = max(0, POMODORO_DURATION - elapsed_time)
        
        if time_left > 0:
            col_info, col_stop = st.columns([3, 1])
            with col_info:
                st.warning(f"🎯 **Sedang fokus pada:** {st.session_state.active_pomodoro_task}")
            with col_stop:
                if st.button("⏹️ Hentikan Sesi", use_container_width=True):
                    st.session_state.pomodoro_running = False
                    st.session_state.active_pomodoro_task = None
                    st.rerun()
            
            progress = elapsed_time / POMODORO_DURATION
            st.progress(min(progress, 1.0))
            
            minutes, seconds = divmod(int(time_left), 60)
            st.metric("⏰ Sisa Waktu", f"{minutes:02d}:{seconds:02d}")
            
            time.sleep(1)
            st.rerun()
        else:
            st.success(f"🎉 **Selamat!** Sesi fokus untuk '{st.session_state.active_pomodoro_task}' telah selesai!")
            st.balloons()
            
            # Refleksi singkat setelah sesi
            st.subheader("📝 Refleksi Singkat")
            reflection = st.text_area("Apa yang Anda pelajari dari sesi ini?", placeholder="Tulis refleksi Anda...")
            if st.button("Simpan Refleksi", use_container_width=True):
                st.success("Refleksi tersimpan! Teruslah berkembang!")
            
            st.session_state.pomodoro_running = False
            st.session_state.active_pomodoro_task = None
            time.sleep(2)
            st.rerun()
    else:
        tasks_to_focus = [t for t in tasks_aktif if t.get('Prioritas') == "Lakukan Sekarang"]
        
        if not tasks_to_focus:
            st.warning("⚠️ Tidak ada tugas dengan prioritas 'Lakukan Sekarang'.")
            st.info("💡 Pergi ke tab 'Prioritaskan' untuk memindahkan tugas ke 'Lakukan Sekarang'.")
        else:
            st.subheader("🎯 Pilih tugas untuk difokuskan:")
            for task in tasks_to_focus:
                with st.container(border=True):
                    col_task, col_btn = st.columns([3, 1])
                    with col_task:
                        area_icon = DEVELOPMENT_AREAS.get(task.get('Area'), {}).get('icon', '📋')
                        st.write(f"{area_icon} **{task.get('Tugas')}**")
                        if task.get('Deskripsi'):
                            st.caption(task.get('Deskripsi', '')[:80])
                        st.caption(f"📅 Deadline: {task.get('Tanggal Selesai Target').strftime('%d %b %Y') if task.get('Tanggal Selesai Target') else '-'}")
                    with col_btn:
                        if st.button("🚀 Mulai Fokus", key=f"focus_{task['id']}", use_container_width=True):
                            st.session_state.pomodoro_running = True
                            st.session_state.active_pomodoro_task = task.get('Tugas')
                            st.session_state.pomodoro_start_time = time.time()
                            st.rerun()

# --- TAB 8: DELEGASI ---
with tab8:
    st.header("🤝 Delegasikan Tugas")
    st.markdown("""
    ### 📋 Panduan Delegasi Efektif untuk Pengembangan Tim
    1. Tentukan tugas yang bisa didelegasikan
    2. Pilih orang yang tepat dengan skill sesuai
    3. Berikan instruksi yang jelas
    4. Tetapkan deadline dan ekspektasi
    5. Lakukan follow-up secara berkala
    
    💡 *Delegasi bukan berarti menghindari tanggung jawab, tapi mengoptimalkan potensi tim!*
    """)
    
    tasks_to_delegate = [t for t in tasks_aktif if t.get('Prioritas') == "Delegasikan"]
    
    if not tasks_to_delegate:
        st.info("📭 Tidak ada tugas yang perlu didelegasikan.")
    else:
        st.warning("⚠️ Tugas berikut direkomendasikan untuk didelegasikan:")
        
        with st.form("delegation_form"):
            delegation_data = {}
            for task in tasks_to_delegate:
                area_icon = DEVELOPMENT_AREAS.get(task.get('Area'), {}).get('icon', '📋')
                st.markdown(f"**{area_icon} {task.get('Tugas')}**")
                delegation_data[task['id']] = st.text_input(
                    "Delegasikan kepada:",
                    value=task.get('Delegasi', ''),
                    placeholder="Nama/Departemen",
                    key=f"delegate_{task['id']}"
                )
                st.markdown("---")
            
            if st.form_submit_button("💾 Simpan Delegasi", use_container_width=True):
                for task_id, delegated_to in delegation_data.items():
                    task = get_task_by_id(task_id)
                    if task and delegated_to:
                        task['Delegasi'] = delegated_to
                        save_task_to_csv(task)
                st.success("✅ Informasi delegasi berhasil disimpan!")
                st.rerun()

# --- TAB 9: LAPORAN ---
with tab9:
    st.header("📈 Laporan & Analisis Pengembangan Diri")
    
    if not st.session_state.tasks:
        st.info("📭 Belum ada data tugas untuk dianalisis.")
    else:
        df_all = pd.DataFrame(st.session_state.tasks)
        
        # Statistik ringkas
        col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)
        with col_stat1:
            total_tasks = len(df_all)
            st.metric("📊 Total Tugas", total_tasks)
        with col_stat2:
            completed_tasks = df_all['Selesai'].sum() if 'Selesai' in df_all else 0
            completion_rate = (completed_tasks / total_tasks * 100) if total_tasks > 0 else 0
            st.metric("✅ Tingkat Penyelesaian", f"{completion_rate:.1f}%")
        with col_stat3:
            active_count = len(tasks_aktif)
            st.metric("🔄 Tugas Aktif", active_count)
        with col_stat4:
            urgent_count = len(tasks_mendesak)
            st.metric("🔥 Tugas Mendesak", urgent_count)
        
        st.markdown("---")
        
        # Laporan 5 area pengembangan diri
        render_development_report()
        
        st.markdown("---")
        
        # Distribusi tugas per area
        st.subheader("📊 Distribusi Tugas per Area Pengembangan Diri")
        
        if 'Area' in df_all.columns:
            area_distribution = df_all[df_all['Area'].notna()]['Area'].value_counts().reset_index()
            area_distribution.columns = ['Area', 'Jumlah']
            
            if not area_distribution.empty:
                area_distribution['Area Name'] = area_distribution['Area'].apply(
                    lambda x: DEVELOPMENT_AREAS.get(x, {}).get('name', x)
                )
                
                fig_area = px.pie(
                    area_distribution,
                    values='Jumlah',
                    names='Area Name',
                    title='Distribusi Tugas Berdasarkan Area',
                    hole=0.3,
                    color_discrete_sequence=px.colors.qualitative.Set3
                )
                st.plotly_chart(fig_area, use_container_width=True)
            else:
                st.info("Belum ada tugas yang dikategorikan ke area pengembangan diri.")
        
        st.markdown("---")
        
        # Grafik tren penyelesaian
        st.subheader("📈 Tren Penyelesaian Tugas (7 Hari Terakhir)")
        
        if 'Tanggal Selesai' in df_all.columns:
            df_all['Tanggal Selesai'] = pd.to_datetime(df_all['Tanggal Selesai'], errors='coerce')
            completed_tasks_df = df_all[df_all['Selesai'] == True].dropna(subset=['Tanggal Selesai'])
            
            if not completed_tasks_df.empty:
                today_date = date.today()
                seven_days_ago = today_date - timedelta(days=6)
                
                recent_completions = completed_tasks_df[
                    (completed_tasks_df['Tanggal Selesai'].dt.date >= seven_days_ago) &
                    (completed_tasks_df['Tanggal Selesai'].dt.date <= today_date)
                ]
                
                tasks_per_day = recent_completions.groupby(recent_completions['Tanggal Selesai'].dt.date).size().reset_index(name='Jumlah')
                tasks_per_day.columns = ['Tanggal Selesai', 'Jumlah']
                
                date_range = pd.date_range(start=seven_days_ago, end=today_date).to_frame(index=False, name='Tanggal Selesai')
                date_range['Tanggal Selesai'] = date_range['Tanggal Selesai'].dt.date
                
                merged_data = pd.merge(date_range, tasks_per_day, on='Tanggal Selesai', how='left').fillna(0)
                
                fig_bar = px.bar(
                    merged_data,
                    x='Tanggal Selesai',
                    y='Jumlah',
                    title='Jumlah Tugas yang Diselesaikan per Hari',
                    labels={'Tanggal Selesai': 'Tanggal', 'Jumlah': 'Jumlah Tugas'},
                    text_auto=True,
                    color_discrete_sequence=['#007BFF']
                )
                fig_bar.update_traces(textposition='outside')
                st.plotly_chart(fig_bar, use_container_width=True)
            else:
                st.info("Belum ada tugas yang diselesaikan dalam 7 hari terakhir.")
        
        st.markdown("---")
        
        # Rekomendasi pengembangan
        st.subheader("💡 Rekomendasi Pengembangan Diri")
        
        dev_data = load_development_data()
        if dev_data:
            recommendations = []
            for area_key, area_info in DEVELOPMENT_AREAS.items():
                progress, _, target = get_area_progress(area_key)
                if progress < 30:
                    recommendations.append(f"🔴 **{area_info['name']}**: Progress masih {progress:.0f}%. Mulai dengan aktivitas sederhana: {area_info['activities'][0]}")
                elif progress < 60:
                    recommendations.append(f"🟡 **{area_info['name']}**: Progress {progress:.0f}%. Tingkatkan konsistensi dengan: {area_info['activities'][1]}")
                elif progress < 80:
                    recommendations.append(f"🟢 **{area_info['name']}**: Progress bagus ({progress:.0f}%). Coba tantangan baru: {area_info['activities'][2]}")
            
            if recommendations:
                for rec in recommendations:
                    st.info(rec)
            else:
                st.success("🎉 Luar biasa! Semua area menunjukkan progress yang baik. Pertahankan konsistensi Anda!")
        else:
            st.info("Mulai isi progress di tab '5 Area Pengembangan Diri' untuk mendapatkan rekomendasi personal.")

# Footer
st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: #666;'>© 2024 Safier Plan - Kembangkan Diri Secara Seimbang di 5 Area | Mental • Sosial • Spiritual • Emosional • Fisik</p>", 
    unsafe_allow_html=True
)